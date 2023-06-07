Attribute VB_Name = "ModGeral"
 
'******************************************************************************************
' Data: 09/03/2003
' Alteração: - Alteração do nome da entidade tblHistoricoContribuicaoMelhoria para
'            tblHistoricoContribuicaoMelhor.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/03/2003
' Alteração: - Alteração do nome da entidade tblHistoricoOcorrenciaPatrimonio para
'            tblHistoricoOcorrenciaPatrimon.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/03/2003
' Alteração: - Alteração do nome da entidade tblMelhoramentoDaSecaoDeLogradouro para
'            tblMelhoramentoDaSecaoDeLograd.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/03/2003
' Alteração: - Alteração do nome da entidade tblMelhoriaContribuicaoMelhoria para
'            tblMelhoriaContribuicaoMelhor.
' Responsável: Everton Bianchini
'******************************************************************************************
    


Option Explicit

    '******************************************************************************************
    ' CONSTANTES UTILIZADAS PARA CLASSE DE INTEGRAÇÃO COM O WORD
    '******************************************************************************************
    
    Public Const WORDOPENFORMATTEMPLATE = 2
    Public Const WORDOPENFORMATDOCUMENT = 1
    Public Const WORDALIGNPARAGRAPHCENTER = 1
    Public Const WORDALIGNPARAGRAPHRIGHT = 2
    Public Const WORDALIGNPARAGRAPHLEFT = 0
    
    '******************************************************************************************
    
    
    '******************************************************************************************
    ' CONSTANTES UTILIZADAS PELO MÓDULO DE PROTOCOLO
    '
    ' OBS: Estas constantes foram transferidas do ModProtocolo para o ModGeral pelo Salsicha
    '      pois o módulo de Contabilidade está tentando utilizar constantes declaradas abaixo.
    '******************************************************************************************
    
    'Botões
    Public Const gstrCalcularReajuste = "CALCULARREAJUSTE"
    Public Const gstrPrecoPublicoGuias = "PRECOPUBLICO"
    Public Const gstrParametroProtocolo = "tblParametroProtocolo"

    Public gblnCadastroProtocolo    As Boolean
    Public gblnTramitacaoProtocolo  As Boolean
    '******************************************************************************************
    
    
    '----- Constante para Pkid fixo no Select do Tag
    Public Const gintPkidFixo       As Integer = 0
    
    '********************************************************************
    ' Constantes para armazenar o tamanho de campos específicos para formatação (zeros a esquerda)
    
    Public Const gintLenInscricao   As Integer = 20
    Public Const gintLenEmissao     As Integer = 3
    Public Const gintLenNumAviso    As Integer = 10
        
    Public vetTamanhoMascaras As TamanhosMascaras
    
    Type TamanhosMascaras
        intMaskImobiliario   As Integer
        intMaskEconomico     As Integer
        intMaskDividaAtiva   As Integer
        intMaskAcordo        As Integer
        intMaskPrecoPublico  As Integer
        intMaskIssConstrucao As Integer
    End Type
    
    'Utilizado na rotina de integração
    Type TabelaIntegracao
        Tabela  As String
        Inseriu As Boolean
    End Type
    
    '********************************************************************
    
    '----- Variáveis para tratamento de banco de dados
    
    'Variaveis de flag para Ordenação de Grid
    Public gblnOrdenacaoAscGrid     As Boolean
    Public gbytOrdenacaoGrid        As Byte
    
    'Variaveis temporarias para Login Unico
    
    Public gstrUsername             As String
    Public gstrPassword             As String
    
    Public gstrBancosDeDados        As String
    Public gcncADOMain              As ADODB.Connection
    Public gcmdADOCmdConMain        As ADODB.Command
    Public gprmADOParamConMain      As ADODB.Parameter
    Public gLocalDoBancoDados       As LocalBanco
    Public gstrServidor             As String
    Public gstrDatabase             As String
    Public gstrCodDataBase          As String
    Public gstrDiretorio            As String
    Public bytLocalBD               As Byte
    Public gblnDBNaoEstaOK          As Boolean
    
    Public conMainADO               As New clsConeccao
    Public gobjBanco                As clsBanco
    
    Public gstrLoginUser            As String
    Public gstrPwdUser              As String
    Public gblnMaster               As Boolean
    Public gblnAdmin                As Boolean
    Public Const bytBancoLocal  As Byte = 0
    Public Const bytBancoRemoto As Byte = 1

    Enum LocalBanco
        mLocal = 0
        mServidor = 1
    End Enum

    '-----------------------------------------------
    'Constantes com os nomes das Views
    '-----------------------------------------------
    Public Const gstrContaValAcumuladosDiario = "vw_ContaValAcumuladosDiario"
    Public Const gstrContaValoresAcumulados = "vw_contavaloresacumulados"
    Public Const gstrContaValAcumuladosEducacao = "vw_contavalacumulados_educacao"
    Public Const gstrContaValAcumuladosIDesp = "VW_CONTAVALACUMULADOS_IDESP"
    Public Const gstrVw_OrcAnual_Realizado = "Vw_OrcAnual_Realizado"
    Public Const gstrVw_OrcAnual = "Vw_OrcAnual"
    Public Const gstrVw_Despesa_Realizada = "Vw_Despesa_Realizada"
    Public Const gstrVw_Ordem_Pagamento = "Vw_Ordem_Pagamento"
    Public Const gstrVwOrdemPagamentoItem = "vwOrdemPagamentoItem"

    '----- Constantes com os nomes das tabelas
    Public Const gstrPORCENTAGEMRECEITAFUTURA = "TBLPORCENTAGEMRECEITAFUTURA"
    Public Const gstrIssEmpresa = "TblIssEmpresa"
    Public Const gstrListaServicoFederal = "tblListaServicoFederal"
    Public Const gstrEspecialidadesSaude = "tblEspecialidadesSaude"
    Public Const gstrParametroAtualizacao = "tblParametroAtualizacao"
    Public Const gstrResumoBancario = "tblResumoBancario"
    Public Const gstrTipoContaBancaria = "tblTipoContaBancaria"
    Public Const gstrConciliacaoTesouraria = "tblConciliacaoTesouraria"
    Public Const gstrConciliacaoExtrato = "tblConciliacaoExtrato"
    Public Const gstrSaldoExtrato = "tblSaldoExtrato"
    Public Const gstrTipoIsencaoImunidade = "tblTipoIsencaoImunidade"
    Public Const gstrCatalogoAssuntoReceita = "tblCatalogoAssuntoReceita"
    Public Const gstrFormaAtualizacaoValor = "tblFormaAtualizacaoValor"
    Public Const gstrReceitasExercicio = "tblReceitasExercicio"
    Public Const gstrMoedas = "tblMoedas"
    Public Const gstrItemEmpenho = "tblItemEmpenho"
    Public Const gstrExecutivoParcela = "tblExecutivoParcela"
    Public Const gstrExecutivoEnvolvidos = "tblExecutivoEnvolvidos"
    Public Const gstrItemEmpenhoAnulado = "tblItemEmpenhoAnulado"
    Public Const gstrRequisicaoComprasSituacoes = "tblRequisicaoComprasSituacoes"
    Public Const gstrTipoRequisicao = "tblTipoRequisicao"
    Public Const gstrTipoMovimentoInventario = "tblTipoMovimentoInventario"
    Public Const gstrCategoriaConstrucao = "tblCategoriaConstrucao"
    Public Const gstrHistoricoPublicidades = "tblHistoricoPublicidades"
    Public Const gstrHistoricoFaceDeQuadra = "tblHistoricoFaceDeQuadra"
    Public Const gstrDocumentosImobiliario = "tblDocumentosImobiliario"
    Public Const gstrFechamentoContabil = "tblFechamentoContabil"
    Public Const gstrTiposDeVias = "tblTiposDeVias"
    Public Const gstrInventarioPatrimonio = "tblInventarioPatrimonio"
    Public Const gstrFaceDeQuadra = "tblFaceDeQuadra"
    Public Const gstrCeps = "tblCEPS"
    Public Const gstrAutorizacaoDeCompra = "tblAutorizacaodeCompra"
    Public Const gstrAutorizacaoDeComprasItens = "tblAutorizacaodeComprasItens"
    Public Const gstrLocalEntrega = "tblLocaisEntrega"
    Public Const gstrAlmoxarifado = "tblAlmoxarifado"
    Public Const gstrOrgaoEntidade = "tblOrgaoEntidade"
    Public Const gstrParcelamentoEntrega = "tblParcelamentoEntrega"
    Public Const gstrComprasLicitacao = "tblComprasLicitacao"
    Public Const gstrLicitacao = "tblLicitacoes"
    Public Const gstrInventarioMaterial = "tblInventarioMaterial"
    Public Const gstrParametrosEspecificos = "tblParametrosEspecificos"
    Public Const gstrRequisicaoMaterial = "tblRequisicaoMaterial"
    Public Const gstrHistoricoAlmoxarifado = " tblHistoricoAlmoxarifado"
    Public Const gstrReservaDotacao = "tblReservaDotacao"
    Public Const gstrReservaDotacaoLiberada = "tblReservaDotacaoLiberada"
    Public Const gstrSubElementoEmpenho = "tblSubElementoEmpenho"
    Public Const gstrSubEmpenhoNF = "tblSubEmpenhoNF"
    Public Const gstrCruzamentoContaExtra = "tblCruzamentoContaExtra"
    Public Const gstrCredorExtra = "tblCredorExtra"

    Public Const gstrTipoCodigoBarra = "tblTipoCodigoBarra"
    Public Const gstrControleAdiantamento = "tblControleAdiantamento"
    Public Const gstrCodigoDeBaixa = "tblCodigoBaixa"
    
    Public Const gstrContaArrecadacaoReceita = "tblContaArrecadacaoReceita"
    Public Const gstrArrecadacaoReceita = "tblArrecadacaoReceita"
    Public Const gstrLancamentoContabil = "tblLancamentoContabil"
    Public Const gstrProcessoPagamento = "tblProcessoPagamento"
    Public Const gstrBolsaDePrecos = "tblBolsaDePreco"
    Public Const gstrRequisicaoCompras = "tblRequisicaoCompras"
    Public Const gstrArquivamentoProcesso = "tblArquivamentoProcesso"
    Public Const gstrLocalArquivamento = "tblLocalArquivamento"
    Public Const gstrJuntada = "tblJuntada"
    Public Const gstrProtocolizacaoVolume = "tblProtocolizacaoVolume"
    Public Const gstrGuiaRemessa = "tblGuiaRemessa"
    Public Const gstrTipoProcesso = "tblTipoProcesso"
    Public Const gstrRegioes = "tblRegioes"
    Public Const gstrImobiliarioProprietarios = "tblImobiliarioProprietarios"
    Public Const gstrDominios = "tblDominios"
    Public Const gstrRespostasPadrao = "tblRespostasPadrao"
    Public Const gstrImagem = "tblImagem"
    Public Const gstrAgencia = "tblAgencia"
    Public Const gstrBairro = "tblBairro"
    Public Const gstrBanco = "tblBanco"
    Public Const gstrCidade = "tblMunicipio"
    Public Const gstrEscolaridade = "tblEscolaridade"
    Public Const gstrEstadoCivil = "tblEstadoCivil"
    Public Const gstrLogradouro = "tblLogradouro"
    Public Const gstrNacionalidade = "tblNacionalidade"
    Public Const gstrNivel = "tblNivel"
    Public Const gstrTipoLogradouro = "tblTipoLogradouro"
    Public Const gstrUF = "tblUF"
    Public Const gstrTituloLogradouro = "tblTituloLogradouro"
    Public Const gstrDiasNaoUteis = "tblDiasNaoUteis"
    Public Const gstrCepsLogradouro = "tblCepsLogradouro"
    Public Const gstrHistoricoLogradouro = "tblHistoricoLogradouro"
    Public Const gstrConfiguracao = "tblConfiguracao"
    Public Const gstrConfiguracaoGeral = "tblConfiguracaoGeral"
    Public Const gstrFormaPagamento = "tblFormaDePagamento"
    Public Const gstrGrupoMaterialServico = "tblGrupoMaterialServico"
    Public Const gstrJustificaJulgamento = "tblJustificaJulgamento"
    Public Const gstrUnidadeMedida = "tblUnidadeMedida"
    Public Const gstrTipoMaterialServico = "tblTipoMaterialServico"
    Public Const gstrLimiteLicitacao = "tblLimiteLicitacao"
    Public Const gstrIndexadorEconomico = "tblIndexadorEconomico"
    Public Const gstrIndiceEconomico = "tblIndiceEconomico"
    Public Const gstrOrdenador = "tblOrdenador"
    Public Const gstrTipoDeComunicacao = "tblTipoDeComunicacao"
    Public Const gstrFormaDeComunicacaoAgencia = "tblFormaDeComunicacaoAgencia"
    Public Const gstrContribuinte = "tblContribuinte"
    Public Const gstrHorarioEspecial = "tblHorarioEspecial"
    Public Const gstrFormaDeComunicacao = "tblFormaDeComunicacao"
    Public Const gstrFormaDeComunicacaoCorretor = "tblFormaDeComunicacaoCorretor"
    Public Const gstrContaBancariaContribuinte = "tblContaBancariaContribuinte"
    Public Const gstrContaBancaria = "tblContaBancaria"
    Public Const gstrSetorFiscal = "tblSetorFiscal"
    Public Const gstrDistritoFiscal = "tblDistritoFiscal"
    Public Const gstrOrgao = "tblOrgao"
    Public Const gstrFuncaoDoGoverno = "tblFuncaoDoGoverno"
    Public Const gstrSubFuncaoGoverno = "tblSubFuncaoGoverno"
    Public Const gstrPrograma = "tblPrograma"
    Public Const gstrProjeto = "tblProjeto"
    Public Const gstrSubPrograma = "tblSubPrograma"
    Public Const gstrProgramaDeTrabalho = "tblProgramaDeTrabalho"
    Public Const gstrElementoDespesa = "tblElementoDespesa"
    Public Const gstrTipoCredito = "tblTipoCredito"
    Public Const gstrVinculo = "tblVinculo"
    Public Const gstrMascaras = "tblMascaras"
    Public Const gstrMascarasItens = "tblMascarasItens"
    Public Const gstrUnidadeCentroDeCusto1 = "tblUnidadeCentroDeCusto1"
    Public Const gstrUnidadeCentroDeCusto2 = "tblUnidadeCentroDeCusto2"
    Public Const gstrPlanoConta = "tblPlanoConta"
    Public Const gstrComLicitacao = "tblComLicitacao"
    Public Const gstrCatalogoMaterialServico = "tblCatalogoMaterialServico"
    Public Const gstrCatalogoMaterialServicoUnid = "tblCatalogoMaterialServicoUnid"
    Public Const gstrUnidadeOrcamentaria = "tblUnidadeOrcamentaria"
    Public Const gstrSubUnidade = "tblSubUnidade"
    Public Const gstrPeriodicidade = "tblPeriodicidade"
    Public Const gstrSeguradoras = "tblSeguradoras"
    Public Const gstrTipoSeguro = "tblTipoSeguro"
    Public Const gstrPeriodo = "tblPeriodo"
    Public Const gstrMarcas = "tblMarcas"
    Public Const gstrNaturezaBens = "tblNaturezaBens"
    Public Const gstrCentrodeCustos = "tblUnidadeCentrodeCusto2"
    Public Const gstrBens = "tblBens"
    Public Const gstrBensComponentes = "tblBensComponentes"
    Public Const gstrContratos = "tblContratos"
    Public Const gstrContratoMaterialServico = "tblContratoMaterialServico"
    Public Const gstrDocumentos = "tblDocumentos"
    Public Const gstrPrevisaoDaReceita = "tblPrevisaoDaReceita"
    Public Const gstrCodigoOrcamentario = "tblCodigoOrcamentario"
    Public Const gstrTipoAssunto = "tblTiposAssuntos"
    Public Const gstrGrupoAssunto = "tblGruposAssuntos"
    Public Const gstrEstantePrateleira = "tblEstantePrateleira"
    Public Const gstrOcorrenciaPatrimonio = "tblOcorrenciaPatrimonio"
'    Public Const gstrHistoricoOcorrenciaPatrimonio = "tblHistoricoOcorrenciaPatrimonio"
    Public Const gstrHistoricoOcorrenciaPatrimonio = "tblHistoricoOcorrenciaPatrimon"
    Public Const gstrHabilitacao = "tblHabilitacao"
    Public Const gstrContribuicaoMelhoria = "tblContribuicaoMelhoria"
    Public Const gstrImobiliario = "tblImobiliario"
    Public Const gstrImobiliarioRural = "tblImobiliarioRural"
    Public Const gstrVeiculosMaquinas = "tblVeiculosMaquinas"
    Public Const gstrDividaAtiva = "tblDividaAtiva"
    Public Const gstrEconomico = "tblEconomico"
    Public Const gstrEstConservacao = "tblEstConservacao"
    Public Const gstrHistoricoValor = "tblHistoricoValor"
    Public Const gstrHistoricoSeguros = "tblHistoricoSeguros"
    Public Const gstrHistoricoCC = "tblHistoricoCentroCusto"
    Public Const gstrTipoCombustivel = "tblTipoCombustivel"
    Public Const gstrFinanceiroVM = "tblFinanceiroVM"
    Public Const gstrPneusVM = "tblPneusVM"
    Public Const gstrSeguroVM = "tblSeguroVM"
    Public Const gstrEventosFrota = "tblEventosFrota"
    Public Const gstrOcorrencias = "tblOcorrencias"
    Public Const gstrTipoImpostosLicencas = "tblTipoImpostosLicencas"
    Public Const gstrComposicaoDaReceita = "tblComposicaoDaReceita"
    Public Const gstrOcorrencia = "tblOcorrencia"
    Public Const gstrCampoDeInscricao = "tblCampoDeInscricao"
    Public Const gstrDetalheDividaAtiva = "tblDetalheDividaAtiva"
    Public Const gstrReceitasDiversas = "tblReceitasDiversas"
    Public Const gstrCaracteristicaDoImovel = "tblCaracteristicasDoImovel"
    Public Const gstrCaracteristicaGeral = "tblCaracteristicaGeral"
    Public Const gstrMelhoramentoPublico = "tblMelhoramentoPublico"
'    Public Const gstrMelhoramentoDaSecaoDeLogradouro = "tblMelhoramentoDaSecaoDeLogradouro"
    Public Const gstrMelhoramentoDaSecaoDeLogradouro = "tblMelhoramentoDaSecaoDeLograd"
'    Public Const gstrMelhoriaContribuicaoMelhoria = "tblMelhoriaContribuicaoMelhoria"
    Public Const gstrMelhoriaContribuicaoMelhoria = "tblMelhoriaContribuicaoMelhor"
    Public Const gstrTabelaDeEdital = "tblTabelaDeEdital"
    Public Const gstrEditalSecaoLogradouro = "tblEditalSecaoLogradouro"
    Public Const gstrAtividadeEC = "tblAtividadeEC"
    Public Const gstrSubGrupoDeAtividade = "tblSubGrupoDeAtividade"
    Public Const gstrGrupoDeAtividade = "tblGrupoDeAtividade"
    Public Const gstrHistoricoContribuinte = "tblHistoricoContribuinte"
    Public Const gstrTipoDeArea = "tblTipoDeArea"
    Public Const gstrTipoDeTestada = "tblTipoDeTestada"
    Public Const gstrSecaoLogradouro = "tblSecaoDeLogradouro"
    Public Const gstrPneus = "tblPneus"
    Public Const gstrPosicaoPneus = "tblPosicaoPneus"
'    Public Const gstrHistoricoContribuicaoMelhoria = "tblHistoricoContribuicaoMelhoria"
    Public Const gstrHistoricoContribuicaoMelhoria = "tblHistoricoContribuicaoMelhor"
    Public Const gstrTributoContribuicaoMelhoria = "tblTributoContribuicaoMelhoria"
    Public Const gstrContador = "tblContador"
    Public Const gstrAtividadeBasica = "tblAtividadeBasica"
    Public Const gstrUtilizacaoDaTabelaDeValor = "tblUtilizacaoDaTabelaDeValor"
    Public Const gstrDetalheDaCaracteristica = "tblDetalheDaCaracteristica"
    Public Const gstrLivrosDeISS = "tblLivrosDeISS"
    Public Const gstrAtividadeDaEmpresa = "tblAtividadeDaEmpresa"
    Public Const gstrTributoEmpresa = "tblTributoEmpresa"
    Public Const gstrValorDaFaixa = "tblValorDaFaixa"
    Public Const gstrTabelaDeValor = "tblTabelaDeValor"
    Public Const gstrValorFaixaEmpresa = "tblValorFaixaEmpresa"
    Public Const gstrFaixaDeValor = "tblFaixaDeValor"
    Public Const gstrSocioEconomico = "tblSocioEconomico"
    Public Const gstrSocio = "tblSocio"
    Public Const gstrHistoricoEconomico = "tblHistoricoEconomico"
    Public Const gstrCaracteristicaDoEconomico = "tblCaracteristicasDoEconomico"
    Public Const gstrCaracteristicaDoImovelRural = "tblCaracteristicaDoImovelRural"
    Public Const gstrValorCompoRec = "tblValorCompoRec"
    Public Const gstrHistoricoImobiliarioRural = "tblHistoricoImobiliarioRural"
    Public Const gstrProducaoImobiliarioRural = "tblProducaoImobiliarioRural"
    Public Const gstrAreaImobiliario = "tblAreaImobiliario"
    Public Const gstrHistoricoImobiliario = "tblHistoricoImobiliario"
    Public Const gstrTestadaImobiliario = "tblTestadaImobiliario"
    Public Const gstrEquipamentoImobiliario = "tblEquipamentoImobiliario"
    Public Const gstrPromissario = "tblPromissario"
    Public Const gstrTextoCarta = "tblTextoCarta"
    Public Const gstrTextoSolicitacao = "tblTextoSolicitacao"
    Public Const gstrTextoAtendimento = "tblTextoAtendimento"
    Public Const gstrDespacho = "tblDespacho"
    Public Const gstrDespachoProtocolo = "tblDespachoProtocolo"
    Public Const gstrReceita = "tblReceita"
    Public Const gstrCatalogoAssunto = "tblCatalogoAssunto"
    Public Const gstrParcelaReceita = "tblParcelaReceita"
    Public Const gstrMensagem = "tblMensagem"
    Public Const gstrParcelaTaxa = "tblParcelaTaxa"
    Public Const gstrObjetivo = "tblObjetivo"
    Public Const gstrEmpenho = "tblEmpenho"
    Public Const gstrFundamentoLegal = "tblFundamentoLegal"
    Public Const gstrRestoaPagar = "tblRestoaPagar"
    Public Const gstrDespesaExtraOrcamentaria = "tblDespesaExtraOrcamentaria"
    Public Const gstrDatabases = "tblDataBase"
    Public Const gstrUsuarios = "tblUsuario"
    Public Const gstrCatalogoTabela = "tblCatalogoTabela"
    Public Const gstrExercicio = "tblExercicio"
    Public Const gstrFonteRecurso = "tblFonteRecurso"
    Public Const gstrGrupoDeFonteRecurso = "tblGrupoDeFonteRecurso"
    Public Const gstrEmpresa = "tblEmpresa"
    Public Const gstrNumeradorDocumentos = "tblNumeradorDocumentos"
    Public Const gstrValorMetroTerreno = "tblValorMetroTerreno"
    Public Const gstrLocaisPorUnidade = "tblLocaisPorUnidade"
    Public Const gstrLocais = "tblLocais"
    Public Const gstrResponsavelPatrimonio = "tblResponsavelPatrimonio"
    Public Const gstrGerarFundef = "tblGerarFundef"
    Public Const gstrUsuariosAlmoxarifado As String = "tblUsuariosAlmoxarifado"
    Public Const gstrTipoPlaqueta = "tblTipoPlaqueta"
    Public Const gstrSituacaoCotacao = "tblSituacaoCotacao"
    Public Const gstrParametroMateriais = "tblParametroMateriais"
    Public Const gstrTipoSolicitacaoCompras = "tblTipoSolicitacaoCompras"
    Public Const gstrPedidoDeEmpenho = "tblPedidoDeEmpenho"
    Public Const gstrPedidoDeEmpenhoItens = "tblPedidoDeEmpenhoItens"
    
    Public Const gstrPlanoContaSaldo = "tblPlanoContaSaldo"
    Public Const gstrEvento = "tblEvento"
    Public Const gstEventoCodigoOrcamentarioCredito = "tblEventoCodigoOrcamentarioCredito"
    Public Const gstEventoCodigoOrcamentarioDebito = "tblEventoCodigoOrcamentarioDebito"
    Public Const gstrEventoContaContabilCredito = "tblEventoContaContabilCredito"
    Public Const gstrEventoContaContabilDebito = "tblEventoContaContabilDebito"
    Public Const gstrEmpenhoContrato = "tblEmpenhoContrato"
    Public Const gstrEmpenhoContratoItens = "tblEmpenhoContratoItens"
    Public Const gstrCartaFianca = "tblCartaFianca"
    Public Const gstrFormaAtualizacao = "tblFormaAtualizacao"
    Public Const gstrSubEmpRetencaoOrcamentaria = "tblSubEmpRetencaoOrcamentaria"
    Public Const gstrmovliq = "tblmovliq"
    Public Const gstrProfissionaisDaSaude = "tblProfissionaisDaSaude"
    Public Const gstrUnidadesDeSaude = "tblUnidadesDeSaude"
    Public Const gstrRegistroDePrecos = "tblRegistroDePrecos"
    Public Const gstrRegistroDePrecosItens = "tblRegistroDePrecosItens"
    
    '--------------------------------------------------
    'Constantes com os nomes das tabelas - Orçamentario
    '--------------------------------------------------
    Public Const gstrTipoLegislacao = "tblTipoLegislacao"
    Public Const gstrTipoConvenio = "tblTipoConvenio"
    Public Const gstrParametrosContabeis = "tblParametrosContabeis"
    Public Const gstrAplicacaoDeDotacao = "tblAplicacaoDeDotacao"
    Public Const gstrParametro = "tblParametro"
    Public Const gstrClassificacaoDespesa = "tblClassificacaoDespesa"
    Public Const gstrSubempenhoLiquidado = "tblSubempenhoLiquidado"
    Public Const gstrAtuacaoLegislacao = "tblAtuacaoLegislacao"
    Public Const gstrModalidade = "tblModalidade"
    Public Const gstrConvenio = "tblConvenio"
    Public Const gstrComplementoEvento = "tblComplementoEvento"
    Public Const gstrDescricaoEmpenho = "tblDescricaoEmpenho"
    Public Const gstrCodigoDiverso = "tblCodigoDiverso"
    Public Const gstrContaUG = "tblContaUG"
    Public Const gstrHistorico = "tblHistorico"
    Public Const gstrTipoReceita = "tblTipoReceita"
    Public Const gstrPlanoPlurianual = "tblPlanoPlurianual"
    Public Const gstrReceitaEstimadaRealizada = "tblReceitaEstimadaRealizada"
    Public Const gstrGrupoDeDespesa = "tblGrupoDeDespesa"
    Public Const gstrOrigem = "tblOrigem"
    Public Const gstrTipoEmpenho = "tblTipoEmpenho"
    Public Const gstrOrigemMaterial = "tblOrigemMaterial"
    Public Const gstrDescentralizacaoCredito = "tblDescentralizacaoCredito"
    Public Const gstrProcessoPagamentoAdiantamento = "tblProcPagtoAdianta"
    Public Const gstrAnulacaoRecPagtoAnulado = "tblAnulacaoRecPgtoAnulado"
    Public Const gstrMovimentacaoDeConvenio = "tblMovimentacaoDeConvenio"
    Public Const gstrCotaFinanceira = "tblCotaFinanceira"
    Public Const gstrCategoria = "tblCategoria"
    Public Const gstrDocumentoContabil = "tblDocumentoContabil"
    Public Const gstrGrupoContaPublica = "tblGrupoContaPublica"
    Public Const gstrProgramacaoDesembolso = "tblProgramacaoDesembolso"
    Public Const gstrPagamentoRegularizar = "tblPagamentoRegularizar"
    Public Const gstrOrdemPagamento = "tblOrdemPagamento"
    Public Const gstrOrdemPagamentoEmpenho = "tblOrdemPagamentoEmpenho"
    Public Const gstrOrdemPagamentoResto = "tblOrdemPagamentoResto"
    Public Const gstrOrdemPagamentoDespesaExtra = "tblOrdemPagamentoDespesaExtra"
    Public Const gstrOrdemPagamentoAnulacaoReceita = "tblOrdemPagamentoAnulacaoRec"
    Public Const gstrControleCheque = "tblControleCheque"
    Public Const gstrMovimentoBanco = "tblMovimentoBanco"
    Public Const gstrEntradaRecurso = "tblEntradaRecurso"
    Public Const gstrMovimentoEconomico = "tblMovimentoEconomico"
    Public Const gstrPlanejamentoOrcamentario = "tblPlanejamentoOrcamentario"
    Public Const gstrCreditoReducao = "tblCreditoReducao"
    Public Const gstrTipoRecurso = "tblTipoRecurso"
    Public Const gstrFundo = "tblFundo"
    Public Const gstrItemDespesa = "tblItemDespesa"
    Public Const gstrComplementoEmpenho = "tblComplementoEmpenho"
    Public Const gstrParcelaRestoPagar = "tblParcelaRestoPagar"
    Public Const gstrParcelaRestoProcessada = "tblParcelaRestoProcessada"
    Public Const gstrPagamento = "tblPagamento"
    Public Const gstrLancamentoConta = "tblLancamentoConta"
    Public Const gstrContaLancamentoContabil = "tblContaLancamentoContabil"
    Public Const gstrProcessoPagtoAnulado = "tblProcessoPagtoAnulado"
    Public Const gstrSubempenhoPagtoAnulado = "tblSubempenhoPagtoAnulado"
    Public Const gstrParcelaRestoPagtoAnulado = "tblParcelaRestoPagtoAnulado"
    Public Const gstrDespesaExtraOrcamPagtoAnulado = "tblDespExtraOrcamPgtoAnulado"
    Public Const gstrSuplementacaoReducaoReceita = "tblSuplementacaoReducaoReceita"
    Public Const gstrReceitaEstimada = "tblReceitaEstimada"
    Public Const gstrDespesaEstimada = "tblDespesaEstimada"
    Public Const gstrCotaTrimestral = "tblCotaTrimestral"
    Public Const gstrProduto = "tblProduto"
    Public Const gstrProgramaEAcao = "tblProgramaEAcao"
    Public Const gstrPropostaAprovada = "tblPropostaAprovada"
    Public Const gstrRestoProcessadoExtra = "tblRestoProcessadoExtra"
    Public Const gstrRestoProcessadoRetencao = "tblRestoProcessadoRetencao"
    Public Const gstrInvoltoraAplicacaoDotacao = "tblInvoltoraAplicacaoDotacao"
    Public Const gstrTransferenciaBancaria = "tblTransferenciaBancaria"
    Public Const gstrContaTransferenciaBancaria = "tblContaTransferenciaBancaria"
    Public Const gstrQuadroSecao = "tblQuadroSecao"
    Public Const gstrFonteDeRecurso = "tblFonteRecurso"
    Public Const gstrAnulacaoDespesa = "tblAnulacaoDespesa"
    Public Const gstrImpressaoFolha = "tblImpressaoFolha"
    Public Const gstrImpressaoCheque = "tblImpressaoCheque"
    Public Const gstrExecutivoMoedas = "TblExecutivoMoedas"
    Public Const gstrExecutivoAdvogados = "tblExecutivoAdvogados"
    Public Const gstrNumeroProtocolo = "tblNumeroProtocolo"
    Public Const gstrPPADespesaAnexoII = "tblPPADespesaAnexoII"
    Public Const gstrPPADespesaAnexoIIMetasCustos = "tblPPADespesaAnexoIIMetasCustos"
    Public Const gstrPPADespesaAnexoIII = "tblPPADespesaAnexoIII"
    Public Const gstrPPADespesaAnexoIIIAcoes = "tblPPADespesaAnexoIIIAcoes"
    Public Const gstrProgFinanceiroDespesa = "tblProgFinanceiroDespesa"

'---- Chave - KEY - dos botões comuns da barra de ferramentas dos sistemas
    Public Const gstrAprovar = "APROVAR"
    Public Const gstrGeraCodigoReduzido = "GERACODIGOREDUZIDO"
    Public Const gstrIncluiElementoDespesa = "INCLUIELEMENTODESPESA"
    Public Const gstrIncluiProjetoAtividade = "INCLUIPROJETOATIVIDADE"
    Public Const gstrCalcular = "CALCULAR"
    Public Const gstrGeraFundef = "GERAFUNDEF"
    Public Const gstrImportarDados = "IMPORTARDADOS"

    'Originalmente do Tributário agora sendo compartilhado.
    Public Const gstrOcorrenciaDoEconomico = "tblOcorrenciaDoEconomico"
    Public Const gstrProcessoEconomico = "tblProcessoEconomico"
    Public Const gstrHistoricoEconVariavel = "tblhistoricoeconvariavel"
    Public Const gstrParametroTributario = "tblParametroTributario"
    Public Const gstrParametrosTributario = "tblParametrosTributario"
    Public Const gstrIsencaoPeriodo = "TblIsencaoPeriodo"
    Public Const gstrIsencaoReceita = "TblIsencaoReceita"
    Public Const gstrIsencaoImunidade = "tblIsencaoImunidade"
    Public Const gstrTributoTipo = "tblTributoTipo"
    Public Const gstrDativa = "tblDativa"
    Public Const gstrDaParcel = "tblDaParcel"
    Public Const gstrLancamentoEconomico = "tbllancamentoeconomico"
    Public Const gstrLancamentoEconIss = "tbllancamentoeconiss"
    Public Const gstrLctEconCaracBoletim = "tbllcteconcaracboletim"
    Public Const gstrLctEconomicoAtividade = "tbllcteconomicoatividade"
    Public Const gstrLctEconomicoTributo = "tbllcteconomicotributo"
    Public Const gstrLctEconPublicidade = "tbllcteconpublicidade"
    Public Const gstrLctEconomicoFeira = "tbllcteconomicofeira"
    Public Const gstrLctEconomicoSocio = "tbllcteconomicosocio"
    Public Const gstrAtivEmpresaTributo = "tblativempresatributo"
    Public Const gstrFaixaPontosPredio = "tblFaixaPontosPredio"
    Public Const gstrExercicioValorM2Predio = "tblExercicioValorM2Predio"
    Public Const gstrLancamentoPagamento = "tblLancamentoPagamento"
    Public Const gstrTributoExercicio = "tblTributoExercicio"
    
    Public Const gstrServico = "TblServico"
    Public Const gstrReferenciasDeTributos = "TblreferenciasDeTributos"
    Public Const gstrListaServicoExercicio = "tblListaServicoExercicio"
    Public Const gstrListaServico = "TbllistaServico"
    Public Const gstrTipoIss = "tblTipoIss"
    Public Const gstrTipoFeira = "tblTipoFeira"
    Public Const gstrFeira = "tblFeira"
    Public Const gstrEconomicoFeira = "tblEconomicoFeira"
    Public Const gstrLoteamento = "tblLoteamento"
    Public Const gstrGuias = "tblGuias"
    Public Const gstrLancamentoAlfa = "tblLancamentoAlfa"
    Public Const gstrLancamentoValor = "tblLancamentoValor"
    Public Const gstrExecutivo = "tblExecutivo"
    Public Const gstrLancamentoGuias = "tbllancamentoguias"
    Public Const gstrLancamentoReceita = "tblLancamentoReceita"
    Public Const gstrLancamentoPPublico = "tblLancamentoPPublico"
    Public Const gstrLancamentoPPublicoReceita = "tblLancamentoPPublicoReceita"
    Public Const gstrTipoFormaCalculo = "tblTipoFormaCalculo"
    Public Const gstrTributo = "tblTributo"
    Public Const gstrTributosFaixa = "tblTributosFaixa"
    Public Const gstrAtividadeTributo = "tblAtividadeTributo"
    Public Const gstrAtividadeTributoTributo = "tblAtividadeTributoTributo"
    
    Public Const gstrAcordoDebitos = "tblAcordoDebitos"
    Public Const gstrResumoTipoPadrao = "tblResumoTipoPadrao"
    Public Const gstrResumoTipoPadraoCarac = "tblResumoTipoPadraoCarac"
    Public Const gstrDebitoAutomatico = "tblDebitoAutomatico"
    
    'Reinato - 16/06/01
    'Trouxe do orcamentario para compartilhar
    'Órgão, Unidade Orçamentária e Sub Unidade
    'no Centro de Custo dos Projetos (Compras, Frotas, Patrimônio, Protocolo, Materiais)
    Public Const gstrUnidadeGestora = "tblUnidadeGestora"
    Public Const gstrGestao = "tblGestao"
    Public Const gstrTipoAdministracao = "tblTipoAdministracao"
    Public Const gstrPoder = "tblPoder"
    Public Const gstrUnidadeFinanceira = "tblUnidadeFinanceira"
    
    'Utilizado Pelo Protocolo/Tributário
    Public Const gstrFormulaDeCalculo = "tblFormulaDeCalculo"
    
    Public Const gstrTiposdeEstocagem = "tblTiposdeEstocagem"
    Public Const gstrMaterialEmEstoque = "tblMaterialEmEstoque"
    Public Const gstrInabilitacao = "tblInabilitacao"

    Public Const gstrPermissoes = "tblPermissao"
    Public Const gstrItens = "tblItem"
    Public Const gstrItemPermissaoEspecifica = "tblItemPermissaoEspecifica"
    Public Const gstrNotaFiscal = "tblNotaFiscal"
    Public Const gstrHistoricoOperacao = "tblHistoricoOperacao"
    Public Const gstrLancamentoCalculo = "tblLancamentoCalculo"
    Public Const gstrProtocolizacaoProcesso = "tblProtocolizacaoProcesso"
    Public Const gstrTramiteProtocolo = "tblTramiteProtocolo"
        
    Public Const gstrBackup = "tblBackup"
    
    'Utilizado Pelo Ouvidoria/Tributário
    Public Const gstrDocumentoEmitido = "tblDocumentoEmitido"
    Public Const gstrReceitaDiversa = "tblReceitaDiversa"
    
    'Utilizado Pelo Patrimonio
    Public Const gstrFechamentoMensal = "tblFechamentoMensal"

    'Material
    Public Const gstrFamiliaMaterial = "tblFamiliaMaterial"
    Public Const gstrModuloContribuinte = "tblModuloContribuinte"
    
    Public Const gstrFechamentoMateriais = "tblFechamentoMateriais"

    'Orcamentario/Compras
    Public Const gstrSuplementacaoReducao = "tblSuplementacaoReducao"
    Public Const gstrDotacaoSuplementadaReduzida = "tblDotacaoSuplementadaReduzida"
    Public Const gstrSuplementacaoReducaoDespesa = "tblSuplementacaoReducaoDespesa"
    Public Const gstrSubempenho = "tblSubempenho"
    Public Const gstrContencaoCredito = "tblContencaoCredito"
    Public Const gstrContencaoCreditoDesbloqueado = "tblContencaoCreditoDesbloquead"
    Public Const gstrPagamentoEstornoEmpenho = "tblPagamentoEstornoEmpenho"
    Public Const gstrCruzamentos = "tblCruzamentos"
    Public Const gstrMovimentoSistemas = "tblMovimentoSistemas"
    Public Const gstrContaMovimentoSistemas = "tblContaMovimentoSistemas"
    
    'Orçamentário/Tributário
    Public Const gstrcheque = "tblcheque"
    Public Const gstrchequeOP = "tblchequeOP"

    'Integracao
    Public Const gstrIntegracaoModulos = "tblIntegracaoModulos"
    Public Const gstrIntegracaoTabelas = "tblIntegracaoTabelas"
    Public Const gstrIntegracaoParametros = "tblIntegracaoParametros"
    Public Const gstrIntegracaoParametrosItens = "tblIntegracaoParametrosItens"
    
    '----- RH - AdGover
    Public Const gstrFuncionario = "tblFuncionario"
    Public Const gstrSecretaria = "tblSecretaria"
    Public Const gstrDepartamento = "tblDepartamento"
    Public Const gstrSecao = "tblSecao"
    Public Const gstrSetor = "tblSetor"
    Public Const gstrSituacao = "tblSituacao"
    Public Const gstrProventoDesconto = "tblProventoDesconto"
    Public Const gstrCargo = "tblCargo"
    Public Const gstrPensionista = "tblPensionista"
    Public Const gstrEvolucaoValor = "tblEvolucaoValor"
    Public Const gstrPagamentos = "tblPagamento"
    
    Public Const gstrAssinaturas = "tblAssinaturas"
    
    Public Const gstrCompraPorConvite = "tblCompraPorConvite"
    Public Const gstrCompraPorTomada = "tblCompraPorTomada"
    Public Const gstrCompraPorConcorrencia = "tblCompraPorConcorrencia"
    Public Const gstrCompraPorDispensa = "tblCompraPorDispensa"
    Public Const gstrCompraInexigibilidade = "tblCompraInexigibilidade"
    Public Const gstrFornecimentosAtividades = "tblFornecimentosAtividades"
    Public Const gstrParcelaContrato = "tblParcelaContrato"
    Public Const gstrAdministracaoContrato = "tblAdministracaoContrato"
    
    Public Const gstrViewComprasLicitacao = "vw_ComprasLicitacao"
    
    '----- Variaveis que irão recebe os nomes das tabela do RH (Antigo ou AdGover)
    Public gstrFuncionarioRH               As String
    Public gstrSecretariaRH                As String
    Public gstrDepartamentoRH              As String
    Public gstrSecaoRH                     As String
    Public gstrSetorRH                     As String
    Public gstrOrgaoRH                     As String
    Public gstrSituacaoRH                  As String
    Public gstrProventoDescontoRH          As String
    Public gstrCargoRH                     As String
    Public gstrPensionistaRH               As String
    Public gstrEvolucaoValorRH             As String
    Public gstrPagamentosRH                As String
    Public gstrBairroRH                    As String
    Public gstrFotoRH                      As String
    
    Public gstrBaseADGover               As String
    Public gstrBaseRH                    As String
    
    '**** Variaveis e Constantes utilizadas para filtro dependendo do menu de chamada ****
    'Constantes que identificam o menu que chamou o formulario
    Public Const gbytMenuCadastro     As Byte = 0
    Public Const gbytMenuProposta     As Byte = 1
    Public Const gbytMenuOrcamento    As Byte = 2
    
    Public gbytMenu                   As Byte
    
    Public gstrErrorInStoredProcedure As String
    '*************************************************************************************
    
    '****************TRIBUTARIO****************
    'Constantes para os tipos de composição
    Public Const TYP_IMOBILIARIA = 1
    Public Const TYP_ECONOMICA = 2
    Public Const TYP_DIVIDA_ATIVA = 3
    Public Const TYP_ACORDO = 4
    Public Const TYP_PRECO_PUBLICO = 5
    Public Const TYP_ISS_CONSTRUCAO = 6
    Public Const TYP_OUTROS = 7
    Public Const TYP_IMOBILIARIO_TAXAS = 8
    Public Const TYP_ISS_MOVIMENTO_GISSONLINE = 9
    
    'Constantes para os tipos de tributos
    Public Const TRIBUTO_TIPO_PUBLICIDADE = 0
    Public Const TRIBUTO_TIPO_FEIRAS = 1
    Public Const TRIBUTO_TIPO_OCUPACAO = 2
    Public Const TRIBUTO_TIPO_OUTROS = 3
    Public Const TRIBUTO_TIPO_HORARIO_ESPECIAL = 4
    
    'Constante que identifica o horario especial
    Public Const CATEGORIA_HORARIO_ESPECIAL = 8
            
    'Constantes com valores da tabela de Referencia de Tributos (tblReferenciaDeTributos)
    Public Const GRUPO_IMOB_TERRENO = 1
    Public Const GRUPO_IMOB_TERRENO_APURADO = 2
    
    Public Const FATOR_TOPOGRAFIA = 1
    Public Const FATOR_PEDOLOGIA = 2
    Public Const FATOR_SITUACAO = 3
    Public Const FATOR_ZONEAMENTO = 4
    Public Const FATOR_DESVIO_FERROVIARIO = 5
    Public Const FATOR_CORREGO = 8
    Public Const FATOR_ACESSIBILIDADE = 11
    Public Const FATOR_SUPERFICIE = 14
    Public Const FATOR_FORMA = 15
    
    Public Const FATOR_GLEBA_APURADO = 6
    Public Const FATOR_PROFUNDIDADE_APURADO = 7
    Public Const FATOR_OBSOLESCENCIA_APURADO = 9
    Public Const FATOR_CORREGO_APURADO = 10
    Public Const FATOR_TESTADA_APURADO = 12
    
    '******************************************
    
    '********************************************************************
    
    ' Constantes para os dias da semana.
    Public Const DOMINGO = 1
    Public Const SEGUNDA = 2
    Public Const TERCA = 3
    Public Const QUARTA = 4
    Public Const QUINTA = 5
    Public Const SEXTA = 6
    Public Const SABADO = 7
    
    '********************************************************************
    
    ' Constantes para os tipos de codigo de barras
    Public Const FEBRABAN = 0
    Public Const FICHA_COMPENSACAO = 1
    
    '********************************************************************
    
'---- Chave - KEY - dos botões comuns da barra de ferramentas dos sistemas
    Public Const gstrSalvar = "SALVAR"
    Public Const gstrLimpar = "LIMPAR"
    Public Const gstrDeletar = "DELETAR"
    Public Const gstrFechar = "FECHAR"
    Public Const gstrSair = "SAIR"
    Public Const gstrNovo = "NOVO"
    Public Const gstrImprimir = "IMPRIMIR"
    Public Const gstrAplicar = "APLICAR"
    Public Const gstrLocalizar = "LOCALIZAR"
    Public Const gstrPreencherLista = "PREENCHERLISTA"
    Public Const gstrConsultar = "CONSULTAR"
    Public Const gstrAtualizar = "ATUALIZAR"
    Public Const gstrGrade = "GRADE"
    Public Const gstrRefresh = "REFRESH"
    Public Const gstrGeraArquivo = "GERAARQUIVO"
    Public Const gstrWord = "WORDPAD"
    Public Const gstrBrasao = "Brasao"
    Public Const gstrLogotipo = "Logotipo"
    Public Const gstrIncluirItem = "IncluirItem"
    Public Const gstrExcluirItem = "ExcluirItem"
    Public Const gstrGuiaDeAcordo = "GUIADEACORDO"
    Public Const gstrGuiaCertidaoNegativa = "GUIACERTIDAONEGATIVA"
    Public Const gstrGuiaCertidaoPositiva = "GUIACERTIDAOPOSITIVA"
    Public Const gstrGuiaCertidaoPositivaEfeitoNegativa = "GUIACERTIDAOPOSITIVAEFEITONEGATIVO"
    Public Const gstrGuiaCertidaoDividaAtiva = "GUIACERTIDAODIVIDAATIVA"
    Public Const gstrGuiaRelacaoDeDebitos = "GUIARELACAODEDEBITOS"
    Public Const gstrParcelamentoDebitoAtualizado = "PARCELAMENTODEBITOATUALIZADO"
    Public Const gstrCancelarReativar = "CANCELARREATIVAR"
    Public Const gstrImprimirGuia = "IMPRIMIRGUIA"
    Public Const gstrProcessamentoBaixa = "PROCESSAMENTOBAIXA"
    Public Const gstrMarcaTudo = "Marca tudo"
    Public Const gstrDesmarcaTudo = "Desmarca tudo"
    Public Const gstrAtualizaDataCombo = "AtualizaDataCombo"
    Public Const gstrCancelar = "CANCELAR"
    Public Const gstrImportar = "IMPORTAR"
    Public Const gstrExportar = "EXPORTAR"
    Public Const gstrEncaminhar = "ENCAMINHAR"
    Public Const gstrEncerrar = "ENCERRAR"
    Public Const gstrConferir = "CONFERIR"
    Public Const gstrReativar = "REATIVAR"
    Public Const gstrModulos = "MODULOS"
    Public Const gstrGerarOrcamento = "GERARORCAMENTO"
    
    Public Const gstrCriarVolume = "Criar Volume"
    Public Const gstrArquivar = "Arquivar"
    Public Const gstrDesarquivar = "Desarquivar"
    Public Const gstrApensar = "Apensar"
    Public Const gstrDesapensar = "Desapensar"
    Public Const gstrTramitar = "Tramitar"
    
    Public Const gstrCriarRemessa = "Criar Remessa"
    Public Const gstrReceberRemessa = "Receber Remessa"
    
    Public Const gstrIncluirGrid As String = "INCLUIRGRID"
    Public Const gstrExcluirGrid As String = "EXCLUIRGRID"
    
    Public Const gstrReserva As String = "RESERVA"
    
'============ Vaariavel usada para Definir titulo da tela de Contribuintes
    Public gstrContribuinteTituloPl As String
    Public gstrContribuinteTituloSg As String
    
'============ Constantes usadas para Definir formatação de campos nas tabelas criadas no word
    Public Const FORMAT_NEGRITO As String = "#BOLD#"

'============ Constantes dos nomes dos bands dos menus
    Public Const gstrMnuArquivo = "MNUARQUIVO"
    Public Const gstrBtnArquivo = "BTNARQUIVO"
    
    Public gintModulo                       As Integer
    
    'Variáveis usadas para controle de parametrização do sistema (Utilizadas no RH)
    Public gblnImprimeTRCTAoGravar          As Boolean
    Public gintEstiloMarcado                As Integer
    
    Public gblnDemonstracao As Boolean
        
    'Variáveis para mascaras específicas
    Public gstrMascaraContaContabil         As String
    Public gstrMascaraCodigoOrcamentario    As String
    Public gstrMascaraElementoDespesa       As String
    Public gstrMascaraItemDespesa           As String
    
    'Variável usada para saber se um form esta sendo chamado de outro.'
    'no Command Button cmd_.... gblnContadorCidade = true             '
    Global gblnContadorCidade      As Boolean
    
    
    '--Variavel usada para a tela frmDataPrompt -----
    Global strDataPrompt As String

    Public gblnRestartRelatorio  As Boolean
    
    Public gintCodSeguranca      As Integer
    Public vetPermissoes() As Permissoes
    
    Type Permissoes
        intCodigo    As Integer
        strPermissao As String
    End Type
    
    Public gblnCancelarInclusao     As Boolean
    
    Public gobjAux              As Object
    Public gblnCancelar         As Boolean
    Public gblnProgressBar  As Boolean
    Public gintMunicipioEmpresa As Integer
'
    Public Const HWND_TOP = 0
    Public Const HWND_BOTTOM = 1
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2
    
    Public Const SWP_SHOWWINDOW = &H40
    Public Const SWP_NOZORDER = &H4
    Public Const SWP_NOACTIVATE = &H10
    Public Const SWP_NOSIZE = &H1
    Public Const SWP_NOMOVE = &H2
    
    Public mouseloc As POINTAPI
    Public retvalue As Boolean

    ' Retorna códigos de registro de funções
    Public Const ERROR_SUCCESS = 0&
    Public Const ERROR_BADDB = 1009&
    Public Const ERROR_BADKEY = 1010&
    Public Const ERROR_CANTOPEN = 1011&
    Public Const ERROR_CANTREAD = 1012&
    Public Const ERROR_CANTWRITE = 1013&
    Public Const ERROR_ACCESS_DENIED = 5&
    Public Const ERROR_OUTOFMEMORY = 14&
    Public Const ERROR_INVALID_PARAMETER = 87&
    
    'Parâmetro de configuração específica para usuário
    Public Const gblnTeclaEnterIgualTab As Boolean = True
    Public gblnMostraDicas              As Boolean
    Public gblnFlagDicas                As Boolean
    Public gblnListagemAutomatica       As Boolean
    Public gblnListViewComGrade         As Boolean
    Public gblnConfirmaGravacao         As Boolean
    Public gblnConfirmaExclusao         As Boolean
    Public gbytchkFundoObjDiferente     As Boolean
    Public gbytRelatorioComEmissor      As Boolean
    Public gblnRetornaRegistro          As Boolean
    Public gblnRelatorioZebrado         As Boolean
    Public gblnObjInacessivelDiferente  As Boolean
    Public gblnSuprimeLinhaEmBranco     As Boolean
    Public gblnCarac_GraficoNoGabarito  As Boolean
    Public gvntCorZebrado               As Variant
    Public gvntFundoObjInacessivel      As Variant
    Public gIntExercicioUsuario         As Integer
    Public gstrFontGabarito             As String
    Public gstrNomeUsuario              As String
    Public gstrDelimitadorPagina        As String * 1
    Public gstrtxtDelimitadorCampo      As String * 1
    Public gintTamanhoDoPapel           As Integer
    Public giContador                   As Integer
    Public gStrSql                      As String
    Public glngCodUsr                   As Long
    Public gstrPKId                     As String
    Public gintExercicio                As Integer
    Public gintExercicioEmQuestao       As Integer
    Public gbytSituacaoExercicio        As Byte
    Public gblnProposta                 As Boolean
    Public gstrDirDocumentos            As String
'------------------------------------------------------------
'Declaração de vetores para rotina de extenso
'Pedro Paulo Simões
'------------------------------------------------------------
    Dim dblMilhar(1 To 6)               As Double
    Dim strMilharSingular(0 To 5)       As String
    Dim strMilharPlural(0 To 5)         As String
    Dim strCentenaMasculino(0 To 8)     As String
    Dim strCentenaFeminino(0 To 8)      As String
    Dim strDezena(0 To 8)               As String
    Dim strUnidadeMasculino(0 To 19)    As String
    Dim strUnidadeFeminino(0 To 19)     As String
    Dim strConstanteEspecifica(0 To 2)  As String
    
'---------------------
    'Predefined Registry Keys used in hKey Argument
    'Public Const REG_SZ = 1
    Public Const REG_OPTION_NON_VOLATILE = 1
    Public Const HKEY_CLASSES_ROOT = &H80000000
    Public Const HKEY_CURRENT_USER = &H80000001
    Public Const HKEY_LOCAL_MACHINE = &H80000002
    Public Const HKEY_USERS = &H80000003
    Public Const HKEY_PERFORMANCE_DATA = &H80000004
    Public Const STANDARD_RIGHTS_ALL = &H1F0000
    Public Const KEY_QUERY_VALUE = &H1
    Public Const KEY_SET_VALUE = &H2
    Public Const KEY_CREATE_SUB_KEY = &H4
    Public Const KEY_ENUMERATE_SUB_KEYS = &H8
    Public Const KEY_NOTIFY = &H10
    Public Const KEY_CREATE_LINK = &H20
    Public Const SYNCHRONIZE = &H100000
    Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL _
                                 Or KEY_QUERY_VALUE _
                                 Or KEY_SET_VALUE _
                                 Or KEY_CREATE_SUB_KEY _
                                 Or KEY_ENUMERATE_SUB_KEYS _
                                 Or KEY_NOTIFY _
                                 Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
                                 
    Public Const VER_PLATFORM_WIN32_NT = 2
    Public Const VER_PLATFORM_WIN32_WINDOWS = 1
    Public Const gcstCorTextInabilitado = &HE0E0E0
    
    Type POINTAPI
        X   As Long
        Y   As Long
    End Type
    
    'Constantes do ListView
    Public Const LVM_FIRST = &H1000
    Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
    Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
    Public Const LVS_EX_CHECKBOXES = &H4

    'Declaração de funções da API usadas no sistema
    
    Declare Function GetKeyState Lib "USER32" (ByVal nVirtKey As Long) As Integer
    Private Declare Function FindWindowEx Lib "USER32" _
    Alias "FindWindowExA" _
    (ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long
    
    Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
    
    Private Declare Function SendTBMessage Lib "USER32" _
    Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Integer, _
    ByVal lParam As Any) As Long
 
     Public Declare Function SendMessageLong Lib "USER32" Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
    Public Declare Sub mouse_event Lib "USER32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
    
    Public Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long

    Public Declare Function SetCursorPos Lib "USER32" (ByVal X As Long, ByVal Y As Long) As Long
    
    Public Declare Function ShowCursor Lib "USER32" (ByVal bShow As Long) As Long
    
    Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
                                 (lpVersionInformation As OSVERSIONINFO) As Long
    Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
                                 (ByVal hKey As Long, ByVal lpSubKey As String, _
                                  ByVal ulOptions As Long, ByVal samDesired As Long, _
                                  phkResult As Long) As Long
                                  
    Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
    
    Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
    
''    Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
''                            (ByVal hKey As Long, _
''                             ByVal lpValueName As String, _
''                             ByVal Reserved As Long, _
''                             ByVal dwType As Long, _
''                             lpData As Any, _
''                             ByVal cbData As Long) As Long

    Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" _
                            (ByVal lpFileName As String, ByVal nBufferLength As Long, _
                             ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

    
''    Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
''                            (ByVal hKey As Long, _
''                             ByVal lpSubKey As String, _
''                             ByVal Reserved As Long, _
''                             ByVal lpClass As String, _
''                             ByVal dwOptions As Long, _
''                             ByVal samDesired As Long, _
''                             lpSecurityAttributes As SECURITY_ATTRIBUTES, _
''                             phkResult As Long, _
''                             lpdwDisposition As Long) As Long
    
    'Declara as chamadas API para Manipulação do Registry
''    Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
    Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
''    Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
''    Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
    Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
    Declare Function RegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
    Declare Function RegDeleteValue Lib "advapi32" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
    Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    
    'Data type Public Constants
    Public Const REG_NONE = 0
    Public Const REG_EXPAND_SZ = 2
    Public Const REG_BINARY = 3
    'Public Const REG_DWORD = 4
    Public Const REG_DWORD_LITTLE_ENDIAN = 4
    Public Const REG_DWORD_BIG_ENDIAN = 5
    Public Const REG_LINK = 6
    Public Const REG_MULTI_SZ = 7
    Public Const REG_RESOURCE_LIST = 8
    Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9
    Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10
    
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    
    Public Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    
    Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion      As Long
        dwMinorVersion      As Long
        dwBuildNumber       As Long
        dwPlatformId        As Long
        szCSDVersion        As String * 128
    End Type
    
    Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
    End Type
    
    Public gstrQueryParamGeral  As String
    Public gstrArrayUF()        As String * 2
    Public gobjGeral            As Object
    Public gcboAux              As ComboBox
    
    Public gblnContrlClick      As Boolean
    Public gstrNomeEmpresa      As String
    Public gstrCidadeEmpresa    As String
    Public gstrUFEmpresa        As String
    Public gintUFEmpresa        As Integer
    Public gblnTrocaUsuario     As Boolean
    
    Public gobjRelatorio        As ActiveReport
    
    'Constantes utilizadas para armazenar Parametro de Reserva de Dotacao
    Public Const RESERVA_AUTORIZACAO As Byte = 1
    Public Const RESERVA_NENHUMA     As Byte = 0
    Public Const RESERVA_REQUISICAO  As Byte = 2
    
    Public bytParametroReserva       As Byte
    
    'Variaveis utilizadas para retornar dados seleciondos no form frmLista
    Public gintCodigoItemLista     As Integer
    Public gstrDescricaoItemLista  As String
    
    Private Const cMaxPath = 1024
    Private Const cMaxFile = 1024
    Private Const sEmpty = ""
    
    Private Type OPENFILENAME
        lStructSize As Long          ' Filled with UDT size
        hWndOwner As Long            ' Tied to Owner
        hInstance As Long            ' Ignored (used only by templates)
        lpstrFilter As String        ' Tied to Filter
        lpstrCustomFilter As String  ' Ignored (exercise for reader)
        nMaxCustFilter As Long       ' Ignored (exercise for reader)
        nFilterIndex As Long         ' Tied to FilterIndex
        lpstrFile As String          ' Tied to FileName
        nMaxFile As Long             ' Handled internally
        lpstrFileTitle As String     ' Tied to FileTitle
        nMaxFileTitle As Long        ' Handled internally
        lpstrInitialDir As String    ' Tied to InitDir
        lpstrTitle As String         ' Tied to DlgTitle
        flags As Long                ' Tied to Flags
        nFileOffset As Integer       ' Ignored (exercise for reader)
        nFileExtension As Integer    ' Ignored (exercise for reader)
        lpstrDefExt As String        ' Tied to DefaultExt
        lCustData As Long            ' Ignored (needed for hooks)
        lpfnHook As Long             ' Ignored (good luck with hooks)
        lpTemplateName As Long       ' Ignored (good luck with templates)
    End Type
    
    Private Declare Function GetOpenFileName Lib "COMDLG32" _
        Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
    Private Declare Function GetSaveFileName Lib "COMDLG32" _
        Alias "GetSaveFileNameA" (file As OPENFILENAME) As Long
    Private Declare Function GetFileTitle Lib "COMDLG32" _
        Alias "GetFileTitleA" (ByVal szFile As String, _
        ByVal szTitle As String, ByVal cbBuf As Long) As Long
    
    Public Enum EOpenFile
        OFN_READONLY = &H1
        OFN_OVERWRITEPROMPT = &H2
        OFN_HIDEREADONLY = &H4
        OFN_NOCHANGEDIR = &H8
        OFN_SHOWHELP = &H10
        OFN_ENABLEHOOK = &H20
        OFN_ENABLETEMPLATE = &H40
        OFN_ENABLETEMPLATEHANDLE = &H80
        OFN_NOVALIDATE = &H100
        OFN_ALLOWMULTISELECT = &H200
        OFN_EXTENSIONDIFFERENT = &H400
        OFN_PATHMUSTEXIST = &H800
        OFN_FILEMUSTEXIST = &H1000
        OFN_CREATEPROMPT = &H2000
        OFN_SHAREAWARE = &H4000
        OFN_NOREADONLYRETURN = &H8000
        OFN_NOTESTFILECREATE = &H10000
        OFN_NONETWORKBUTTON = &H20000
        OFN_NOLONGNAMES = &H40000
        OFN_EXPLORER = &H80000
        OFN_NODEREFERENCELINKS = &H100000
        OFN_LONGNAMES = &H200000
    End Enum

'-------------------------------------------------------------------------------------------
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                             (ByVal hWnd As Long, _
                              ByVal lpOperation As String, _
                              ByVal lpFile As String, _
                              ByVal lpParameters As String, _
                              ByVal lpDirectory As String, _
                              ByVal nShowCmd As Long) As Long

'Internet Explorer
Private Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long

Public Enum EShellShowConstants
     essSW_HIDE = 0
     essSW_MAXIMIZE = 3
     essSW_MINIMIZE = 6
     essSW_SHOWMAXIMIZED = 3
     essSW_SHOWMINIMIZED = 2
     essSW_SHOWNORMAL = 1
     essSW_SHOWNOACTIVATE = 4
     essSW_SHOWNA = 8
     essSW_SHOWMINNOACTIVE = 7
     essSW_SHOWDEFAULT = 10
     essSW_RESTORE = 9
     essSW_SHOW = 5
End Enum

Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5            '  access denied
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2                     '  file not found
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_PNF = 3                     '  path not found
Private Const SE_ERR_OOM = 8                     '  out of memory
Private Const SE_ERR_SHARE = 26

Global frmAtivo As Form

    Public Enum IconesErro
        Critical = 1
        Exclamation = 2
        Information = 3
        Interrogation = 4
    End Enum

' Usado para AlwaysOnTop
Const flags = 3
Public SetTop As Boolean

Private acbToolBar As ActiveBar2LibraryCtl.tool



'******************************************************************************************
' Declarações referentes à adaptação do software ao Banco de Dados Oracle
'******************************************************************************************
'
' Enumerador utilizado para listagem dos Bancos de Dados com os quais a aplicação é
' compatível.
Public Enum EDatabases
    SQLServer = 1
    Oracle = 2
End Enum

' Variável utilizada para determinar, no software inteiro, o tipo de Banco de Dados utilizado
Public bytDBType As EDatabases

' Variáveis utilizadas para armazenar os comandos nativos do Banco de Dados
Public strGETDATE As String
'Public strISNULL As String
Public strSUBSTRING As String
Public strLen As String
Public strCONCAT As String
Public strOUTJSQLServer As String
Public strOUTJOracle As String
Public strREADPAST As String

Public strROWNUM As String

' Variáveis utilizadas para armazenar formatos
Public strYEAR As String
Public strDAY As String
Public strMONTH As String

' Enumerador utilizado para listagem dos DataTypes com os quais a função gstrCONVERT
' está adaptda. Caso seja necessário utilizar um novo datatype, este deverá ser adicionado
' no enumerador e o seu tratamento na função gstrCONVERT.
Public Enum EConvertDataTypes
    CDT_INT = 1
    CDT_DATETIME = 2
    CDT_VARCHAR = 3
    CDT_NVARCHAR = 4
    cdt_numeric = 5
End Enum

Private strINT As String
Private strVARCHAR As String
Private strDATETIME As String
Private strNVARCHAR As String
Private strNUMERIC As String

' Enumerador utilizado para listagem dos formatos de retorno da função gstrDATEDIFF.
' Caso seja necessário utilizar um novo datatype, este deverá ser adicionado
' no enumerador e o seu tratamento na função gstrDATEDIFF.
Public Enum EDateDiffFormats
    DDF_DAYS = 1
    DDF_MINUTES = 2
End Enum

Private Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As String
  cch As Long
End Type

Const MF_CHECKED = &H8&
Const MF_APPEND = &H100&
Const TPM_LEFTALIGN = &H0&
Const MF_DISABLED = &H2&
Const MF_GRAYED = &H1&
Const MF_SEPARATOR = &H800&
Const MF_STRING = &H0&
Const TPM_RETURNCMD = &H100&
Const TPM_RIGHTBUTTON = &H2&

Public lngMenu As Long

Public Type Volumes
     Pkid       As Long
     DTMDATA    As Date
     strSumula  As String
     Tramitar   As Boolean
     Juntar     As Boolean
     Arquivar   As Boolean
End Type

'Variaveis utilizadas para controlar a primeira situaçao salva
Public EMCAD As Integer
Public Const SIGLA_EMCADASTRAMENTO = "Em cadastramento"

Public GridDeImpressao As TDBGrid

Private intOrientacao       As Integer

Private Declare Function CreatePopupMenu Lib "USER32" () As Long
Private Declare Function TrackPopupMenuEx Lib "USER32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal hWnd As Long, ByVal lptpm As Any) As Long
Private Declare Function AppendMenu Lib "USER32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function DestroyMenu Lib "USER32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "USER32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpmii As MENUITEMINFO) As Long

Public intTamanhoGrupo          As Integer
Public intTamanhoTipo           As Integer
Public intTamanhoFamilia        As Integer
Public intTamanhoItem           As Integer
Public bitCodigoCompleto        As Byte


Public gstrTitulolPedidoCotacao As String
Public gstrFornecedor           As String
Public gstrIdentificacao        As String
Public gstrEndereco             As String
Public gstrTelefone             As String
Public gstrContato              As String
Public gstrObservacao           As String
Public gintLinhaInicial         As Integer
Public gstrSenha                As String
Public gintColNumeroDoItem      As Integer
Public gintColDescricaoDoItem   As Integer
Public gintColComplementoDoItem As Integer
Public gintColMarca             As Integer
Public gintColUnidadeDeMedida   As Integer
Public gintColQuantidade        As Integer
Public gintColValorUnitario     As Integer
Public gintColValorTotal        As Integer
'Variavel Usado para Passar a Tabela como parametro para o Relatorio Externo DocAutorFornecimento
'Ítalo Siqueira
Public gstrTabAutForn           As String
'VARIÁVEIS USADAS NO MÓDULO DE COMPRAS PARA INTERAGIR COM O clsRelatorio
Public Table_Marcas()                 As String
Public Qtd_Marcas                     As Long
Public Modalidade As Boolean 'VARIAVEVL USADA NA FUNÇÃO STRQUERYRELATORIO DO FORMA FRMDOCAUTORFORNECIMENTO

' Variaveis utilizadas para identificar a Categoria da Construção

Type CategoriaConstrucao
    ResidencialHorizontal      As Long
    ResidencialVertical        As Long
    ComercialHorizontal        As Long
    ComercialVertical          As Long
    Industrial                 As Long
    ImobiliarioGeral           As Long
    ImobiliarioTerreno         As Long
    EconomicoGeral             As Long
End Type

Public vetCategoriaConstrucao As CategoriaConstrucao

Public gintRecebimento As Integer

Private Function gstrGetTableAlias(strSql As String, strTabela As String) As String

'******************************************************************************************
' Data: 14/03/2003
' Descrição: - strSQL --> instrução SQL
'            - strTabela --> tabela que deve ter seu apelido pesquisado
' Alteração: - Implementação da função, a qual tem a função de retornar o apelido de uma
'            determinada tabela em uma instrução SQL.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strFrom As String
    Dim strApelido As String
    Dim intInd As Integer
    
    Dim vetReservedWords() As String
    Dim vetFrom() As String
    Dim intIndVetReservedWords As Integer
    
    vetReservedWords = Split("LEFT,RIGHT,INNER,WHERE,GROUP,HAVING,ORDER,DISTINCT", ",")
    
    intInd = InStrRev(UCase(strSql), "FROM")
    
    strApelido = strTabela
    
    If (intInd > 0) Then
        strFrom = Trim(Mid(strSql, intInd + 4))
        
        Do Until InStr(1, strFrom, "  ") = 0
            strFrom = Replace(strFrom, "  ", " ")
        Loop
        
        vetFrom = Split(strFrom, " ")
        
        For intInd = 0 To UBound(vetFrom())
            
            If (InStr(1, UCase(vetFrom(intInd)), UCase(strTabela)) > 0) Then
            
                If (intInd < UBound(vetFrom())) Then
                    
                    strApelido = vetFrom(intInd + IIf((InStr(1, vetFrom(intInd), ",") > 0), 0, 1))
                    If (UCase(strApelido) = "AS") Then
                        strApelido = vetFrom(intInd + 2)
                    End If
                    
                    intInd = InStr(1, strApelido, ",")
                    
                    If (intInd > 0) Then
                        strApelido = Mid(strApelido, 1, intInd - 1)
                    End If
                
                    For intIndVetReservedWords = 0 To UBound(vetReservedWords())
                        If (UCase(strApelido) = vetReservedWords(intIndVetReservedWords)) Then
                            strApelido = strTabela
                            Exit For
                        End If
                        
                    Next intIndVetReservedWords
                    
                End If
            
                Exit For
            
            End If
            
        Next intInd
        
    End If
    
    gstrGetTableAlias = strApelido

End Function
Public Function gstrDATEADD(strFormat As String, strQuantity As String, strDate As String) As String

'******************************************************************************************
' Data: 11/03/2003
' Descrição: - strFormat --> formato da adição. Valores permitidos: DD, MM, YYYY
'            - strQuantity --> quantidade a ser adicinada
'            - strDate --> campo que será adicionado
' Alteração: - Implementação da função, a qual tem a função de retornar o comando nativo
'            adequado, conforme o DB corrente, para adicionar uma quantidade (strQuantity)
'            a um determinado campo (strDate) no formato determinado (strFormat).
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strQttAux As String

    Select Case bytDBType
    
        Case EDatabases.SQLServer
            gstrDATEADD = " (DATEADD(" & strFormat & ", (" & strQuantity & "), " & strDate & ")) "
        
        Case EDatabases.Oracle
        
            Select Case strFormat
            
                Case "DD"
                    gstrDATEADD = " (" & strDate & " + (" & strQuantity & ")) "
                
                Case Else
                    
                    strQttAux = strQuantity
                    
                    If strFormat = "YYYY" Then
                        strQttAux = strQttAux & " * 12"
                    End If
                    
                    gstrDATEADD = " (ADD_MONTHS(" & strDate & ", (" & strQttAux & "))) "
            
            End Select
    
    End Select

End Function

Public Function gstrDATEDIFF(strInitialDate As String, strFinalDate As String, Optional bytFormat As EDateDiffFormats = DDF_DAYS) As String

'******************************************************************************************
' Data: 11/03/2003
' Descrição: - strInitialDate --> Data inicial do intervalo
'            - strFinalDate --> Data final do intervalo
' Alteração: - Implementação da função, a qual tem a função de retornar o comando nativo
'            adequado, conforme o DB corrente, para retornar a diferença em dias entre
'            duas datas.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 08/04/2003
' Alteração: - Adicionado o parâmetro opcional bytFormat, o qual indica o valor a ser
'            retornado pela função. O retorno default é em dias.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strFormat As String

    Select Case bytDBType
    
        Case EDatabases.SQLServer
            
            Select Case bytFormat
                Case EDateDiffFormats.DDF_DAYS
                    strFormat = "DD"
                
                Case EDateDiffFormats.DDF_MINUTES
                    strFormat = "n"
            
            End Select
        
            gstrDATEDIFF = " (DATEDIFF(" & strFormat & ", " & strInitialDate & ", " & strFinalDate & ")) "
        
        Case EDatabases.Oracle
            gstrDATEDIFF = " (TRUNC("
            
            Select Case bytFormat
                Case EDateDiffFormats.DDF_DAYS
                    strFormat = ""
                
                Case EDateDiffFormats.DDF_MINUTES
                    strFormat = " * 24 * 60"
            
            End Select
            
            gstrDATEDIFF = gstrDATEDIFF & _
                "(" & strFinalDate & " - " & strInitialDate & ")" & strFormat
                
            gstrDATEDIFF = gstrDATEDIFF & ", 0)) "
    
    End Select

End Function


Public Function gstrFormataDataOracle(strData As String, Optional strFormatoOracle As String = "yyyy/mm/dd hh24:mi:ss") As String

'******************************************************************************************
' Data: 07/03/2003
' Descrição: - strData --> Data a ser inserida no Banco de Dados
'            - strFormatoOracle --> Formato no qual a data está sendo inserida no Banco de
'                                   Dados. Valor default: yyyy/mm/dd hh24:mi:ss
' Alteração: - Criada a função gstrFormataDataOracle, a qual tem a função de retornar um
'            string  com o comando Oracle que converta a data em uma data válida para o
'            Oracle.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strDataAux As String
    Dim strFormatoVB As String
    
    If (strData = "") Then
        gstrFormataDataOracle = " null "
        Exit Function
    End If
    
    strFormatoVB = Replace(strFormatoOracle, "24", "")
    strFormatoVB = Replace(strFormatoVB, "mi", "nn")
    
    strDataAux = Format(strData, strFormatoVB)
    
    gstrFormataDataOracle = " TO_DATE('" & _
        strDataAux & "', '" & _
        strFormatoOracle & "')"

End Function
Public Function gstrCASEWHEN(strField As String, _
                             strCases As String, _
                             Optional strDefaultCase As String = "") As String

'******************************************************************************************
' Data: 06/03/2003
' Descrição: - strField --> campo a ser analisado na estrutura CASE
'            - strCases --> string contendo valores DE/PARA a serem utilizados na estrutura
'                           CASE. Deve ser passado no formato 'valor de teste,novo valor',
'                           sendo possível n pares de valores.
'            - strDefaultCase --> valor default a ser retornado caso nenhum case seja
'                                 satisfeito.
' Alteração: - Implementação da função, a qual tem a função de criar uma estrutura CASE
'            adequada ao Banco de Dados corrente.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim vetStrCases() As String
    Dim intCtd As Integer
    
    Dim strFirstField As String
    Dim intOperator As Integer
    Dim blnGreaterLessOperators As Boolean
    Dim blnEqualOperator As Boolean
    
    If Trim(strCases) = "" Then Exit Function
    
    intOperator = InStr(1, strField, ">", vbTextCompare)
    intOperator = IIf((intOperator = 0), InStr(1, strField, "<", vbTextCompare), intOperator)
    intOperator = IIf((intOperator = 0), InStr(1, strField, "=", vbTextCompare), intOperator)
    
    If (intOperator > 0) Then
    
        blnGreaterLessOperators = True
        
        If (bytDBType = EDatabases.Oracle) Then
            strFirstField = Trim(Mid(strField, 1, intOperator - 1))
            
            If (InStr(1, strField, ">", vbTextCompare) > 0) Then
                strField = Replace(strField, ">=", ",", , , vbTextCompare)
                strField = Replace(strField, ">", ",", , , vbTextCompare)
                strField = " GREATEST(" & strField & ") "
            
            ElseIf (InStr(1, strField, "<", vbTextCompare) > 0) Then
                strField = Replace(strField, "<=", ",", , , vbTextCompare)
                strField = Replace(strField, "<", ",", , , vbTextCompare)
                strField = " LEAST(" & strField & ") "
            
            Else
                blnEqualOperator = True
                strField = Replace(strField, "=", ",", , , vbTextCompare)
            
            End If
            
        End If
        
    End If
    
    If bytDBType = EDatabases.SQLServer Then
        gstrCASEWHEN = " CASE "
    ElseIf bytDBType = EDatabases.Oracle Then
        gstrCASEWHEN = " DECODE("
    End If
    
    If (bytDBType = EDatabases.SQLServer) And (blnGreaterLessOperators) Then
        gstrCASEWHEN = gstrCASEWHEN & " WHEN "
    
    End If
    
    gstrCASEWHEN = gstrCASEWHEN & strField
    
    If ((UCase(Mid(strCases, 1, 6)) <> " CASE ") And (bytDBType = EDatabases.SQLServer)) Or _
        ((UCase(Mid(strCases, 1, 8)) <> " DECODE(") And (bytDBType = EDatabases.Oracle)) Then
    
        If (bytDBType = EDatabases.Oracle) And (blnGreaterLessOperators) Then
            strCases = strFirstField & "," & strCases
        End If
        
'        vetStrCases = Split(strCases, ",", , vbTextCompare)
        vetStrCases = strarrSplit(strCases)
        
        If (bytDBType = EDatabases.SQLServer) And (blnGreaterLessOperators) Then
            gstrCASEWHEN = gstrCASEWHEN & " THEN "
            gstrCASEWHEN = gstrCASEWHEN & vetStrCases(intCtd)
        
        Else
            For intCtd = 0 To UBound(vetStrCases()) Step 2
                If bytDBType = EDatabases.SQLServer Then
                    gstrCASEWHEN = gstrCASEWHEN & " WHEN "
                    gstrCASEWHEN = gstrCASEWHEN & vetStrCases(intCtd)
                    gstrCASEWHEN = gstrCASEWHEN & " THEN "
                    gstrCASEWHEN = gstrCASEWHEN & vetStrCases(intCtd + 1)
                ElseIf bytDBType = EDatabases.Oracle Then
                    gstrCASEWHEN = gstrCASEWHEN & ", "
                    gstrCASEWHEN = gstrCASEWHEN & vetStrCases(intCtd)
                    gstrCASEWHEN = gstrCASEWHEN & ", "
                    gstrCASEWHEN = gstrCASEWHEN & vetStrCases(intCtd + 1)
                End If
            Next
        
        End If
    
    Else
    
        If (bytDBType = EDatabases.SQLServer) Then
            gstrCASEWHEN = gstrCASEWHEN & " THEN "
            
        ElseIf (bytDBType = EDatabases.Oracle) Then
            gstrCASEWHEN = gstrCASEWHEN & ", "
            
            If (Not blnEqualOperator) Then
                gstrCASEWHEN = gstrCASEWHEN & strFirstField & ", "
            End If
        
        End If
        
        gstrCASEWHEN = gstrCASEWHEN & " (" & strCases & ") "
        
    End If
    
    If Trim(strDefaultCase) <> "" Then
        If bytDBType = EDatabases.SQLServer Then
            gstrCASEWHEN = gstrCASEWHEN & " ELSE "
        ElseIf bytDBType = EDatabases.Oracle Then
            gstrCASEWHEN = gstrCASEWHEN & ", "
        End If
        gstrCASEWHEN = gstrCASEWHEN & strDefaultCase
    End If

    If bytDBType = EDatabases.SQLServer Then
        gstrCASEWHEN = gstrCASEWHEN & " END "
    ElseIf bytDBType = EDatabases.Oracle Then
        gstrCASEWHEN = gstrCASEWHEN & ") "
    End If

End Function


Public Function strarrSplit(ByVal strTextToSplit As String)

'******************************************************************************************
' Data: 12/05/2003
' Descrição: - strTextToSplit --> conjunto de parâmetros a serem separados
' Alteração: - Implementação da função strarrSplit, a qual tem a função de que separar em
'            um array os parâmetros que irão ser passados para uma stored procedure. Esta
'            função assemelha-se com a função, nativa do Visual Basic, Split, porém ela
'            separa corretamente cadeias de caracteres (conjunto de caracteres delimitados
'            por duas aspas simples), comandos nativos do banco de dados etc.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strarrParameters() As String
Dim intInd As Integer
Dim blnAspasSimples As Boolean
Dim blnAbreChaves As Boolean

Dim strCloseBlock As String

Dim strarrCommands() As String

Dim strParameter As String

Dim intIndArrCmds As Integer

strarrCommands = Split(UCase("TO_DATE,TO_CHAR,TO_NUMBER,NVL,ISNULL"), ",", , vbTextCompare)
ReDim strarrParameters(0)
intInd = 1

Do Until strTextToSplit = ""

    blnAspasSimples = (Mid(strTextToSplit, intInd, 1) = "'")
    blnAbreChaves = (Mid(strTextToSplit, intInd, 1) = "{")
    
    If (Not blnAspasSimples) And (Not blnAbreChaves) Then
    
        If (Mid(strTextToSplit, intInd, 1) = ",") Or (intInd > Len(strTextToSplit)) Then
        
            strParameter = LTrim(strTextToSplit)
        
            For intIndArrCmds = LBound(strarrCommands()) To UBound(strarrCommands())
                If UCase(Left(strParameter, Len(strarrCommands(intIndArrCmds)))) = strarrCommands(intIndArrCmds) Then
                    intInd = InStr(1, strParameter, ")", vbTextCompare)
                    intInd = InStr(intInd, strParameter, ",", vbTextCompare)
                    
                    If intInd = 0 Then
                        intInd = Len(strParameter) + 1
                    End If
                    
                    strParameter = Trim(Mid(strTextToSplit, 1, intInd - 1))
                    
                    Exit For
                
                End If
            Next intIndArrCmds
        
            If (intIndArrCmds > UBound(strarrCommands())) Then
                strParameter = Trim(Mid(strTextToSplit, 1, intInd - 1))
            End If
        
            If Trim(strarrParameters(UBound(strarrParameters()))) <> "" Then
                ReDim Preserve strarrParameters(UBound(strarrParameters()) + 1)
            End If
            
            strarrParameters(UBound(strarrParameters())) = strParameter
            
            strTextToSplit = LTrim(Mid(strTextToSplit, intInd + 1))
            intInd = 0
            
        End If
    
    Else
        intInd = InStr(intInd + 1, strTextToSplit, IIf((blnAspasSimples), "'", "}"), vbTextCompare)
    
    End If

    intInd = intInd + 1

Loop

strarrSplit = strarrParameters()

End Function

Public Function gstrCONVERT(bytDatatype As EConvertDataTypes, strField As String) As String

'******************************************************************************************
' Data: 05/03/2003
' Descrição: - bytDatatype --> tipo para o qual o campo deve ser convertido
'            - strField --> campo a ser convertido
' Alteração: - Implementação da função, a qual tem a função de retornar o comando nativo
'            adequado, conforme o DB corrente, para converter um campo em um determinado
'            tipo.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim LStrDatatype As String

    Select Case bytDatatype
        Case EConvertDataTypes.CDT_INT
            LStrDatatype = strINT
        Case EConvertDataTypes.CDT_DATETIME
            LStrDatatype = strDATETIME
        Case EConvertDataTypes.CDT_VARCHAR
            LStrDatatype = strVARCHAR
        Case EConvertDataTypes.CDT_NVARCHAR
            LStrDatatype = strNVARCHAR
        Case EConvertDataTypes.cdt_numeric
            LStrDatatype = strNUMERIC
    End Select

    Select Case bytDBType
        Case EDatabases.SQLServer
            
'            Select Case bytDatatype
            
'                Case EConvertDataTypes.CDT_VARCHAR
'                    gstrCONVERT = " STR(" & strField & ") "
                
'                Case Else
                    gstrCONVERT = " CONVERT (" & LStrDatatype & ", " & strField & ") "
            
'            End Select
        
        Case EDatabases.Oracle
            
            Select Case bytDatatype
                Case EConvertDataTypes.CDT_INT, EConvertDataTypes.cdt_numeric
                    gstrCONVERT = " TO_NUMBER(" & strField & ") "
                
                Case EConvertDataTypes.CDT_DATETIME
                    gstrCONVERT = " CAST (" & strField & " AS " & LStrDatatype & ") "
                
                Case EConvertDataTypes.CDT_NVARCHAR, EConvertDataTypes.CDT_VARCHAR
                    gstrCONVERT = " TO_CHAR(" & strField & ") "
            
            End Select
    
    End Select

End Function

Public Function gstrRIGHT(strField As String, intLenght As Variant) As String
    
    Select Case bytDBType
        Case EDatabases.SQLServer
            
            gstrRIGHT = " RIGHT (" & strField & ", " & intLenght & ") "
            
        Case EDatabases.Oracle
            
            gstrRIGHT = " SUBSTR (" & strField & ", " & intLenght * -1 & ") "
    
    End Select

End Function

Public Function gstrDATEPART(strDateFormat As String, _
                             strDateField As String) As String

'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Implementação da função, a qual tem a função de retornar o comando nativo
'            adequado, conforme o DB corrente, para recuperar uma parte (strDateFormat) de
'            um campo data (strDateField).
' Responsável: Everton Bianchini
'******************************************************************************************

    Select Case bytDBType
    
        Case EDatabases.SQLServer
            Select Case strDateFormat
                
                Case strYEAR
                    gstrDATEPART = " YEAR(" & strDateField & ") "
            
                Case strMONTH
                    gstrDATEPART = " MONTH(" & strDateField & ") "
            
                Case strDAY
                    gstrDATEPART = " DAY(" & strDateField & ") "
            
                Case Else
                    gstrDATEPART = " DATEPART(" & strDateFormat & ", " & strDateField & ")"
            
            End Select
        
        Case EDatabases.Oracle
            gstrDATEPART = " TO_CHAR(" & strDateField & ", '" & strDateFormat & "')"
            
            Select Case strDateFormat
                
                Case strYEAR, strMONTH, strDAY
                
                    gstrDATEPART = " TO_NUMBER(" & gstrDATEPART & ") "
            
            End Select
        
    End Select

End Function


Public Function gstrExtenso(vntNumero As Variant, _
                   Optional bytTipoDeMoeda As Byte, _
                   Optional bytFeminino As Byte) As String
    '-------------------------------------------------------------------
    ' FUNÇÃO USADA PARA RETORNAR UM VALOR POR EXTENSO
    '-------------------------------------------------------------------
    ' PARÂMETRO
    ' 1 - vntNumero - NÚMERO QUE SERÁ EXCRITO POR EXTENSO
    ' 2 - bytTipoDeMoeda - 0 = REAL, 1 = DOLAR, 3 SEM MOEDA
    ' 3 - bytFeminino - 1 = RETONA PALAVRA FEMININA (UMA, DUZENTAS ETC.
    '-------------------------------------------------------------------
    Dim vntPartesDoNumero   As Variant
    Dim strFinal            As String
    Dim strNomeDaMoeda      As String
    Dim strNomeCentezimal   As String
    IniciaConstanteDeExtenso
    vntPartesDoNumero = Split(Format$(vntNumero, "####################0.00"), ",")
    If bytTipoDeMoeda = 0 Then
        If vntPartesDoNumero(0) > 1 Then
            strNomeDaMoeda = "reais "
        Else
            strNomeDaMoeda = "real "
        End If
    ElseIf bytTipoDeMoeda = 1 Then
        If vntPartesDoNumero(0) > 1 Then
            strNomeDaMoeda = "dólar "
        Else
            strNomeDaMoeda = "dólares "
        End If
    End If
    If vntPartesDoNumero(1) > 1 Then
        strNomeCentezimal = "centavos"
    Else
        strNomeCentezimal = "centavo"
    End If
    strFinal = strFinal & gstrConstanteConcatenada(vntPartesDoNumero(0), strNomeDaMoeda, _
                                                   bytFeminino, vntPartesDoNumero(1))
    If vntPartesDoNumero(1) <> 0 Then
        If vntPartesDoNumero(0) = 0 Then
            strFinal = strFinal & gstrConstanteConcatenada(vntPartesDoNumero(1), strNomeCentezimal, bytFeminino)
        Else
            strFinal = strFinal & strConstanteEspecifica(0)
            strFinal = strFinal & gstrConstanteConcatenada(vntPartesDoNumero(1), strNomeCentezimal, bytFeminino)
        End If
    End If
    If Trim(strFinal) = "" Then
        strFinal = "Zero " & strNomeDaMoeda
    Else
        Mid(strFinal, 1, 1) = UCase(Mid(strFinal, 1, 1))
    End If
    gstrExtenso = Trim(strFinal)
End Function

Public Function gstrISNULL(ByVal strField As String, ByVal strTruePart As String, Optional ByVal strFalsePart As String = vbNullString) As String

'******************************************************************************************
' Data: 25/03/2003
' Descrição: - strField --> coluna a ser avaliada
'            - strTruePart --> retorno caso a coluna avaliada seja null
'            - strFalsePart --> retorno caso a coluna avaliada não seja null. Caso seja
'                               omitido então passará a ter o mesmo valor de strField
' Alteração: - Implementação da função, a qual tem a função de retornar o comando nativo
'            adequado, conforme o DB corrente, para verificar se uma coluna (strField) é
'            NULL e retornar um valor caso o seja (strTruePart) ou retornar outro valor
'            caso não seja NULL (strFalsePart). A não utilização da função ISNULL do
'            SQL Server (NVL no Oracle) se deve ao fato de que a concatenação de um valor
'            NULL com um outro valor no SQL Server o retorno é um valor NULL, enquanto que
'            no Oracle o retorno seria o outro valor.
' Responsável: Everton Bianchini
'******************************************************************************************

    If (Trim(strField) = "") Or (Trim(strTruePart) = "") Then
        Exit Function
    End If

    strFalsePart = IIf((Trim(strFalsePart) = ""), strField, strFalsePart)

    If (strField = strFalsePart) Then
    
        Select Case bytDBType
            
            Case EDatabases.SQLServer
                gstrISNULL = " ISNULL("
            
            Case EDatabases.Oracle
            
                gstrISNULL = " NVL("
            
        End Select
        
        gstrISNULL = gstrISNULL & strField & ", " & strTruePart & ") "
    
    Else
        
        Select Case bytDBType
        
            Case EDatabases.SQLServer
                gstrISNULL = gstrCASEWHEN("", "(" & strField & ") IS NULL," & strTruePart, strFalsePart)
            
            Case EDatabases.Oracle
                gstrISNULL = gstrCASEWHEN(strField, "NULL," & strTruePart, strFalsePart)
        
        End Select
    
    End If

End Function

Public Function gstrStoredProcedure(strStoreProcedure As String, _
                                    Optional strParameters As String = vbNullString, _
                                    Optional blnReturnResultset As Boolean = False, _
                                    Optional lngResultSet As Long = 10000) As String

'******************************************************************************************
' Data: 12/03/2003
' Descrição: - strStoreProcedure --> nome da Stored Procedure (SP) a ser executada
'            - strParameters --> parâmetros para execução da SP
'            - blnReturnResultset --> indica se a SP retorna um Resultset. Default FALSE
' Alteração: - Implementação da função, a qual tem a função de retornar uma string de
'            execução de stored procedures adequada ao Banco de Dados corrente.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim recParameters As ADODB.Recordset

Dim strSql As String
Dim strPackageName As String

Dim bytFirstPosition As Byte
Dim strRESULTSET As String

Dim blnREFCURSOR As Boolean

    If (bytDBType = EDatabases.SQLServer) Then
        gstrStoredProcedure = strStoreProcedure & " " & strParameters
    
    ElseIf (bytDBType = EDatabases.Oracle) Then
        gstrStoredProcedure = "{ call "
        
        strStoreProcedure = Trim(strStoreProcedure)
        
        If (blnReturnResultset) Then
            strPackageName = "pe_" & Mid(strStoreProcedure, 4)
       
            strSql = "SELECT ARGUMENT_NAME PARAMETER_NAME, POSITION, DATA_TYPE "
            strSql = strSql & "FROM ALL_ARGUMENTS "
            strSql = strSql & "WHERE (LOWER(PACKAGE_NAME) = '" & LCase(strPackageName) & "') AND "
            strSql = strSql & "(NOT (ARGUMENT_NAME IS NULL)) AND "
            strSql = strSql & "(DATA_TYPE IN ('PL/SQL TABLE', 'REF CURSOR')) AND "
            strSql = strSql & "(IN_OUT = 'OUT') AND "
            strSql = strSql & "(OWNER = 'CPDMASTER') "
            strSql = strSql & "ORDER BY POSITION"
        
            Set gobjBanco = New clsBanco
            Call gobjBanco.CriaADO(strSql, 5, recParameters)
        
            strPackageName = strPackageName & "."
        
        End If
        
        gstrStoredProcedure = gstrStoredProcedure & strPackageName & strStoreProcedure
        
        If (Not (recParameters Is Nothing)) Then
            
            If (Not recParameters.EOF) Then
                
                bytFirstPosition = recParameters("POSITION")
                
                If (UCase(recParameters("DATA_TYPE")) <> "REF CURSOR") Then
                    strRESULTSET = " {RESULTSET " & lngResultSet
'                Else
'                    blnREFCURSOR = True
'
'                End If
                    
                Do Until recParameters.EOF
                
                    strRESULTSET = strRESULTSET & _
                        ", " & recParameters("PARAMETER_NAME")
                        
                    recParameters.MoveNext
                    
                Loop
                    
'                If Not blnREFCURSOR Then
                    strRESULTSET = strRESULTSET & "}"
                
                End If
                
                If Left(strRESULTSET, 1) = "," Then
                    strRESULTSET = Mid(strRESULTSET, 2)
                End If
                
                If (bytFirstPosition = 1) And (Len(strRESULTSET) > 0) Then
                    strParameters = strRESULTSET & IIf((Trim(strParameters) = ""), "", ", " & strParameters)
                
                ElseIf (Len(strRESULTSET) > 0) Then
                    strParameters = IIf((Trim(strParameters) = ""), "", strParameters & ", ") & strRESULTSET
                
                End If
            
            End If
        
        End If
        
        gstrStoredProcedure = gstrStoredProcedure & "(" & strParameters & ") }"
        
    End If

Set recParameters = Nothing

End Function


Public Function gstrTOPnSQLServer(ByVal lngN As Long) As String

'******************************************************************************************
' Data: 25/03/2003
' Descrição: - lngN --> quantidade de registros a ser retornado
' Alteração: - Implementação da função, a qual tem a função de retornar o comando nativo
'            adequado ao SQL Server para que a query retorne somente n registros. Caso o
'            DB corrente seja o Oracle, a função retorna uma string vazia ("") devido a
'            incompatibilidade de localização do comando na query. No SQL Server há ainda
'            a opção do comando TOP n ser acompanhado da palavra PERCENT, o que indicaria
'            que n é uma porcentagem do número total de linhas retornadas pela query. Esta
'            opção não foi implementada pois o Oracle não tem a mesma.
' Localização da execução: SELECT gstrTOPnSQLServer(n)
' Resultado da execução:   SELECT TOP n
' Responsável: Everton Bianchini
'******************************************************************************************

    If (bytDBType = EDatabases.SQLServer) Then
        gstrTOPnSQLServer = " TOP " & CStr(lngN) & " "
    End If

End Function
Public Function gstrTOPnOracle(ByVal strCmdSQL As String, _
                               ByVal lngN As Long, _
                               Optional strCampoFiltro As String, _
                               Optional strValorFiltro As String, _
                               Optional strCampoRetorno As String) As String

'******************************************************************************************
' Data: 25/03/2003
' Descrição: - lngN --> quantidade de registros a ser retornado
'            - strCmdSQL --> comando SQL completo que se tornará uma Top-N query
' Alteração: - Implementação da função, a qual tem a função de retornar o comando nativo
'            adequado ao Oracle para que a query retorne somente n registros. Caso o DB
'            corrente seja o SQL Server, a função retorna o próprio comando SQL passado
'            como parâmetro.
' Localização da execução: strSQL = gstrTOPnOracle(n, strSQL)
' Resultado da execução:   strSQL = "SELECT * FROM (strSQL) WHERE ROWNUM <= n"
' Responsável: Everton Bianchini
'******************************************************************************************

    If (bytDBType = EDatabases.Oracle) Then
        If strCampoFiltro <> "" Then
            
            strCmdSQL = Replace(strCmdSQL, strCampoFiltro & " = " & strValorFiltro & " AND ", "")
            strCmdSQL = Replace(strCmdSQL, " AND " & strCampoFiltro & " = " & strValorFiltro, "")
            strCmdSQL = Replace(strCmdSQL, " " & strCampoFiltro & " = " & strValorFiltro, "")
            strCmdSQL = Replace(strCmdSQL, "WHERE ORDER ", "ORDER ")
            If Right(UCase(Trim(strCmdSQL)), 5) = "WHERE" Or Right(UCase(Trim(strCmdSQL)), 6) = "WHERE)" Then strCmdSQL = Replace(strCmdSQL, "WHERE", "")
            
            strCmdSQL = Left(strCmdSQL, 8) & strCampoFiltro & ", " & Right(strCmdSQL, Len(strCmdSQL) - 8)
            gstrTOPnOracle = " SELECT " & strCampoRetorno & " FROM (" & strCmdSQL & ") WHERE ROWNUM <= " & CStr(lngN) & " AND " & strCampoFiltro & " = " & strValorFiltro
        Else
            gstrTOPnOracle = " SELECT * FROM (" & strCmdSQL & ") WHERE ROWNUM <= " & CStr(lngN)
        End If
    Else
        gstrTOPnOracle = strCmdSQL
    End If

End Function
Public Sub IniciaConstanteDeExtenso()
    '-------------------------------------------------------------------
    ' SUB USADA PARA INICIALIZAR AS CONSTANTES PARA A ROTINA DE EXTENSO
    '-------------------------------------------------------------------
    Dim bytInd  As Byte
    dblMilhar(1) = 1000#
    dblMilhar(2) = 1000000#
    dblMilhar(3) = 1000000000#
    dblMilhar(4) = 1000000000000#
    dblMilhar(5) = 1E+15
    dblMilhar(6) = 1E+18

    strConstanteEspecifica(0) = "e "
    strConstanteEspecifica(1) = "de "
    strConstanteEspecifica(2) = ", "
    
    strMilharSingular(0) = "mil "
    strMilharSingular(1) = "milhão "
    strMilharSingular(2) = "bilhão "
    strMilharSingular(3) = "trilhão "
    strMilharSingular(4) = "quadrilhão "
    strMilharSingular(5) = "quinquilhão "

    strMilharPlural(0) = "mil "
    strMilharPlural(1) = "milhões "
    strMilharPlural(2) = "bilhões "
    strMilharPlural(3) = "trilhões "
    strMilharPlural(4) = "quadrilhões "
    strMilharPlural(5) = "quinquilhões "

    strCentenaMasculino(0) = "cento "
    strCentenaMasculino(1) = "duzentos "
    strCentenaMasculino(2) = "trezentos "
    strCentenaMasculino(3) = "quatrocentos "
    strCentenaMasculino(4) = "quinhentos "
    strCentenaMasculino(5) = "seiscentos "
    strCentenaMasculino(6) = "setecentos "
    strCentenaMasculino(7) = "oitocentos "
    strCentenaMasculino(8) = "novecentos "

    strCentenaFeminino(0) = "cento "
    strCentenaFeminino(1) = "duzentas "
    strCentenaFeminino(2) = "trezentas "
    strCentenaFeminino(3) = "quatrocentas "
    strCentenaFeminino(4) = "quinhentas "
    strCentenaFeminino(5) = "seiscentas "
    strCentenaFeminino(6) = "setecentas "
    strCentenaFeminino(7) = "oitocentas "
    strCentenaFeminino(8) = "novecentas "

    strDezena(0) = "dez "
    strDezena(1) = "vinte "
    strDezena(2) = "trinta "
    strDezena(3) = "quarenta "
    strDezena(4) = "cinquenta "
    strDezena(5) = "sessenta "
    strDezena(6) = "setenta "
    strDezena(7) = "oitenta "
    strDezena(8) = "noventa "
                    
    strUnidadeMasculino(0) = "zero "
    strUnidadeMasculino(1) = "um "
    strUnidadeMasculino(2) = "dois "
    strUnidadeMasculino(3) = "três "
    strUnidadeMasculino(4) = "quatro "
    strUnidadeMasculino(5) = "cinco "
    strUnidadeMasculino(6) = "seis "
    strUnidadeMasculino(7) = "sete "
    strUnidadeMasculino(8) = "oito "
    strUnidadeMasculino(9) = "nove "
    strUnidadeMasculino(10) = "dez "
    strUnidadeMasculino(11) = "onze "
    strUnidadeMasculino(12) = "doze "
    strUnidadeMasculino(13) = "treze "
    strUnidadeMasculino(14) = "quatorze "
    strUnidadeMasculino(15) = "quinze "
    strUnidadeMasculino(16) = "dezesseis "
    strUnidadeMasculino(17) = "dezesete "
    strUnidadeMasculino(18) = "dezoito "
    strUnidadeMasculino(19) = "dezenove "

    For bytInd = 0 To 19
        If bytInd = 1 Then
            strUnidadeFeminino(bytInd) = "uma "
        ElseIf bytInd = 2 Then
            strUnidadeFeminino(bytInd) = "duas "
        Else
            strUnidadeFeminino(bytInd) = strUnidadeMasculino(bytInd)
        End If
    Next
End Sub

Public Sub IniciaVarsCmdsNativosDB()

'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Implementação da função, a qual tem a função de preencher as variáveis de
'            comandos nativos com os seus respectivos comandos, conforme o DB corrente.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Select Case bytDBType
        
        Case EDatabases.SQLServer
            strGETDATE = " GETDATE() "
'            strISNULL = " ISNULL"
            strSUBSTRING = " SUBSTRING"
            strLen = " LEN"
            strCONCAT = " + "
            strOUTJSQLServer = "*"
            strOUTJOracle = ""
            strREADPAST = " WITH (READPAST) "
            
            strDAY = "DD"
            strMONTH = "MM"
            strYEAR = "YYYY"
            
            strINT = "int"
            strVARCHAR = "VARCHAR"
            strDATETIME = "DATETIME"
            strNVARCHAR = "NVARCHAR"
            strNUMERIC = "NUMERIC"
            strROWNUM = ""
            
        Case EDatabases.Oracle
            strGETDATE = " SYSDATE "
'            strISNULL = " NVL"
            strSUBSTRING = " SUBSTR"
            strLen = " LENGTH"
            strCONCAT = " || "
            strOUTJSQLServer = ""
            strOUTJOracle = " (+) "
            strREADPAST = ""
            
            strDAY = "DD"
            strMONTH = "MM"
            strYEAR = "YYYY"
            
            strINT = "NUMBER"
            strVARCHAR = "VARCHAR2"
            strDATETIME = "DATE"
            strNVARCHAR = "VARCHAR2"
            strNUMERIC = "NUMBER"
            
            strROWNUM = "ROWNUM"
            
    End Select

End Sub

Public Sub VerificaCentena(ByVal lngValor As Long, _
                                 strExtenso As String, _
                                 bytFeminino As Byte, _
                                 vntValorDecimal As Variant, _
                                 vntNumero As Variant)
    '-------------------------------------------------------------------
    ' SUB USADA PARA RETORNAR AS CENTENAS, DEZENAS E UNIDADES
    '-------------------------------------------------------------------
    ' PARÂMETRO
    ' 1 - strExtenso - VALOR ESCRITO POR EXTENSO
    ' 2 - lngValor - NÚMERO QUE SERÁ EXCRITO POR EXTENSO
    ' 3 - bytFeminino - 1 = RETONA PALAVRA FEMININA (UMA, DUZENTAS ETC.
    '-------------------------------------------------------------------
    Dim strExtensoAux       As String
    Dim lngValorOriginal    As Long
    lngValorOriginal = lngValor
    If lngValor = 100 Then
        strExtensoAux = strExtensoAux & "cem "
    ElseIf lngValor > 100 Then
        If bytFeminino Then
            strExtensoAux = strExtensoAux & strCentenaFeminino(Int(lngValor / 100) - 1)
        Else
            strExtensoAux = strExtensoAux & strCentenaMasculino(Int(lngValor / 100) - 1)
        End If
    End If
    lngValor = lngValor - ((Int(lngValor / 100)) * 100)
    If lngValor > 19 Then
        If strExtensoAux <> "" Then
            strExtensoAux = strExtensoAux & strConstanteEspecifica(0)
        End If
        strExtensoAux = strExtensoAux & strDezena(Int(lngValor / 10) - 1)
        lngValor = lngValor - (Int(lngValor / 10) * 10)
    End If
    If lngValor <> 0 Then
        If strExtensoAux <> "" Then
            strExtensoAux = strExtensoAux & strConstanteEspecifica(0)
        End If
        If bytFeminino Then
            strExtensoAux = strExtensoAux & strUnidadeFeminino(lngValor)
        Else
            strExtensoAux = strExtensoAux & strUnidadeMasculino(lngValor)
        End If
    End If
    If strExtenso <> "" Then
        If (lngValorOriginal < 100 Or lngValorOriginal Mod 100 = 0) And _
           (vntValorDecimal = 0 Or vntNumero < 20) Then
            strExtensoAux = strConstanteEspecifica(0) & strExtensoAux
        Else
            strExtensoAux = strConstanteEspecifica(2) & strExtensoAux
            strExtenso = Trim(strExtenso)
        End If
    End If
    strExtenso = strExtenso & strExtensoAux
End Sub

Public Function gstrConstanteConcatenada(ByVal vntNumero As Variant, _
                                               strNomeDaMoeda As String, _
                                               bytFeminino As Byte, _
                                      Optional vntDecimal As Variant) As String
    '-------------------------------------------------------------------
    ' FUNÇÃO USADA PARA CONCATENAR AS CONSTANTES PARA O EXTENSO
    '-------------------------------------------------------------------
    ' PARÂMETRO
    ' 1 - vntNumero - NÚMERO QUE SERÁ EXCRITO POR EXTENSO
    ' 2 - bytTipoDeMoeda - 0 = REAL, 1 = DOLAR, 3 SEM MOEDA
    ' 3 - bytFeminino - 1 = RETONA PALAVRA FEMININA (UMA, DUZENTAS ETC.
    '-------------------------------------------------------------------
    Dim intInd              As Integer
    Dim lngValorInteiro     As Long
    Dim vntValorOriginal    As Variant
    Dim strExtenso          As String
    If vntNumero <> 0 Then
        vntValorOriginal = vntNumero
        For intInd = 6 To 1 Step -1
            If vntNumero >= dblMilhar(intInd) Then
                lngValorInteiro = Int(vntNumero / dblMilhar(intInd))
                If lngValorInteiro = 1 Then
                    If strExtenso <> "" Then
                        strExtenso = strExtenso & strConstanteEspecifica(0)
                    End If
                    strExtenso = strExtenso & strUnidadeMasculino(1)
                    strExtenso = strExtenso & strMilharSingular(intInd - 1)
                Else
                    VerificaCentena lngValorInteiro, strExtenso, bytFeminino, vntDecimal, vntNumero
                    strExtenso = strExtenso & strMilharPlural(intInd - 1)
                End If
                vntNumero = vntNumero - (lngValorInteiro * dblMilhar(intInd))
            End If
        Next
        If vntNumero > 0 Then
            VerificaCentena vntNumero, strExtenso, bytFeminino, vntDecimal, vntNumero
        End If
    End If
    If (strNomeDaMoeda <> "") And (strExtenso <> "") Then
        For intInd = 6 To 2 Step -1
            If vntValorOriginal >= dblMilhar(intInd) Then
                lngValorInteiro = Int(vntValorOriginal / dblMilhar(intInd))
                vntValorOriginal = vntValorOriginal - (lngValorInteiro * dblMilhar(intInd))
                If vntValorOriginal = 0 Then
                    strExtenso = strExtenso & strConstanteEspecifica(1)
                    Exit For
                End If
            End If
        Next
        strExtenso = strExtenso & strNomeDaMoeda
    End If
    gstrConstanteConcatenada = strExtenso
End Function

Sub VerificaListaAutomatica(Optional strTabela As String, _
                            Optional objLista As Object, _
                            Optional strQuery As String)
    
    '-------------------------------------------------------------
    ' SUB USADA PARA VERIFICAR O FLAG QUE INDICA SE PREENCHE
    ' O LISTVIEW DOS FOUMULÁRIOS DE CADASTRO AUTOMATICAMENTE
    ' AO CARREGAR O FORMULÁRIO E/OU GRAVAR NOVO REGISTRO
    '-------------------------------------------------------------
    ' PARÂMETRO
    ' 1 - strTabela(tabela que será lida)
    ' 2 - objLista(objeto a ser preenchido)
    ' 3 - strQuery(query específica se não se tratar de um select
    '              simples de uma única tabela
    '-------------------------------------------------------------
                            
    If gblnListagemAutomatica Then
        LeDaTabelaParaObj strTabela, objLista, strQuery
    End If
    
End Sub

'''''Reinato
    
Public Function gvntValorParaObjLista(adoCampo As ADODB.Field)

    Dim blnFlagEspecifico   As Boolean
    
    '-------------------------------------------------------------------
    ' FUNÇÃO USADA PARA VERIFICAR O TIPO DO CAMPO 'adoCampo' E FORMATAR
    ' O VALOR DE SAIDA PARA PREENCHER A LISTA
    '-------------------------------------------------------------------
    ' PARÂMETROS:
    '
    ' 1 - adoCampo(Campo com as informações lidas do SQL)
    '-------------------------------------------------------------------
    
    If InStr("CGC,CPF", UCase(Mid(adoCampo.Name, 4, 3))) <> 0 _
    Or UCase(Mid(adoCampo.Name, 4, 4)) = "CNPJ" Then
        gvntValorParaObjLista = gstrCGCCPFFormatado(gstrENulo(adoCampo))
    ElseIf InStr("CEP", UCase(Mid(adoCampo.Name, 4, 3))) <> 0 Then
        gvntValorParaObjLista = gstrCEPFormatado(gstrENulo(adoCampo))
    ElseIf adoCampo.Type = adNumeric Then
        With adoCampo
            gvntValorParaObjLista = gstrConvVrDoSql(gstrENulo(.Value), .NumericScale, .Precision)
        End With
    ElseIf adoCampo.Type = adCurrency Then
        gvntValorParaObjLista = gstrConvVrDoSql(gstrENulo(adoCampo.Value))
    ElseIf adoCampo.Type = adBoolean Then
        gvntValorParaObjLista = gstrSimOuNao(adoCampo)
    ElseIf adoCampo.Type = adUnsignedTinyInt Then
        If UCase(adoCampo.Name) = "BLNNATUREZADACONTA" Then
            gvntValorParaObjLista = gstrNaturezaDaConta(adoCampo)
        Else
            gvntValorParaObjLista = gstrENulo(adoCampo.Value)
        End If
    ElseIf adoCampo.Type = adDate Or adoCampo.Type = adDBDate Or adoCampo.Type = adDBTimeStamp Then
        gvntValorParaObjLista = gstrDataFormatada(adoCampo)
    ElseIf adoCampo.Type = adVarWChar Or adoCampo.Type = adVarChar Then
        gvntValorParaObjLista = gvntFormatacaoEspecifica(adoCampo)
    Else
        gvntValorParaObjLista = gstrENulo(adoCampo.Value)
    End If
End Function

Public Function gstrCEPFormatado(vntCepAux As Variant) As String
    '--------------------------------------------------------------'
    ' FUNÇÃO USADA PARA FORMATAR O VALOR DO CEP INFORMADO.         '
    '--------------------------------------------------------------'
    ' PARÂMETRO:                                                   '
    '                                                              '
    ' 1 - vntCepAux(Valor digitado)                                '
    '--------------------------------------------------------------'
    If IsNumeric(vntCepAux) Then
        If Val(vntCepAux) > 0 Then
            If Len(Trim(vntCepAux)) = 7 Then
                vntCepAux = "0" & Trim(vntCepAux)
            ElseIf Len(Trim(vntCepAux)) < 8 Then
                vntCepAux = vntCepAux & String$(8 - Len(Trim$(vntCepAux)), "0")
            End If
            gstrCEPFormatado = Format(vntCepAux, "00000\-000")
        Else
            gstrCEPFormatado = ""
        End If
    Else
        gstrCEPFormatado = gstrENulo(vntCepAux)
    End If
End Function

Public Function gstrNaturezaDaConta(vntNatureza As Variant) As String
    '---------------------------------------------------------------------
    ' FUNÇÃO USADA PARA RETORNAR A NATUREZA DA CONTA (Débito ou Crédito)
    '---------------------------------------------------------------------
    ' PARÂMETRO:
    '
    ' 1 - vntNatureza
    '--------------------------------------------------------------
    If vntNatureza = 0 Then
        gstrNaturezaDaConta = "Crédito"
    Else
        gstrNaturezaDaConta = "Débito"
    End If
End Function

Public Sub MarcaDesmarcaTrueDBGrid(intFlag As Integer, _
                                    objLista As Object, _
                                    xadbLista As XArrayDB)
    Dim intInd As Integer
    For intInd = 0 To xadbLista.Count(1) - 1
        xadbLista(intInd, 1) = intFlag
        objLista.Update
    Next
    Set objLista.Array = xadbLista
    objLista.ReBind
    objLista.Refresh
    objLista.Update
End Sub

Public Sub ImprimeRelatorio(objRelatorio As Object, _
                            strQuery As String, _
                   Optional strTitulo As String, _
                   Optional lngIntervaloDeTempo As Long)
    '-------------------------------------------------------------------
    ' SUB USADA PARA VISUALIZACAO DE RELATÓRIOS.
    '-------------------------------------------------------------------
    ' PARÂMETROS:
    '
    ' 1 - objRelatorio - Relatorio
    ' 2 - strQuery - Instrucao SQL passada ao relatorio - Tipo String
    ' 4 - lngIntervaloDeTempo - Indica quanto tempo,em segundo, esperar
    '                           enquanto executa a query
    '-------------------------------------------------------------------

'******************************************************************************************
' Data: 06/03/2003
' Alteração: - Preenchimento da string de conexão conforme o Banco de Dados corrente.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim adoRelatorio            As ADODB.Recordset
    Dim strRelatorio            As String
    Dim strConnectionString     As String
    Dim objActiveReports        As ActiveReport
    Dim blnExisteArquivo        As Boolean
    Dim objField As Object
    
    On Error GoTo ErroImprimeRelatorio
    
    If Trim(strQuery) = "" Then
        Exit Sub
    End If
    If lngIntervaloDeTempo = 0 Then
        lngIntervaloDeTempo = 30
    End If
    Screen.MousePointer = vbHourglass
    
    strRelatorio = objRelatorio.Name & ".rpx"
    
    gblnRestartRelatorio = False
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strQuery, lngIntervaloDeTempo, adoRelatorio) Then
'        objRelatorio.adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrLoginUser & ";pwd=" & gstrPwdUser & ";"
        
        If adoRelatorio.EOF Then
            ExibeMensagem "Nenhum registro encontrado nestas especificações."
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        If bytDBType = EDatabases.SQLServer Then
            strConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
        Else
            strConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
        End If
        
        On Error GoTo NaoExiste
        If Dir(gstrDirDocumentos & "Documentos\Relatorios\" & strRelatorio, vbArchive) <> "" Then
            blnExisteArquivo = True
        Else
NaoExiste:
            blnExisteArquivo = False
        End If
        
        On Error GoTo ErroImprimeRelatorio

        If blnExisteArquivo Then
            Set objActiveReports = New ActiveReport
            objActiveReports.LoadLayout gstrDirDocumentos & "Documentos\Relatorios\" & strRelatorio
            
            objActiveReports.adoDataControl.Provider = ""
            objActiveReports.adoDataControl.ConnectionString = strConnectionString
            objActiveReports.adoDataControl.Source = strQuery
            Set objActiveReports.adoDataControl.Recordset = adoRelatorio
            objActiveReports.adoDataControl.ConnectionTimeout = lngIntervaloDeTempo
            objActiveReports.adoDataControl.CommandTimeout = lngIntervaloDeTempo
            
            frmVisualizarRelatorio.ARViewer.ReportSource = objActiveReports
            If Trim(strTitulo) <> "" Then
                frmVisualizarRelatorio.Caption = strTitulo
            End If
            
            objActiveReports.ResetScripts
            
            objActiveReports.AddCode AdicionaCodigo
            
            objActiveReports.AddNamedItem "clsRelatorio", New clsRelatorio
            
            'Rotina criada por Tiago em 28/12/2004
            
            If blnSubRelatorio(objRelatorio.Name) Then
            
               Dim adoResultado As New ADODB.Recordset
               Dim strSql As String
               
               strSql = "SELECT strSubRelatorio FROM tblSubRelatorios WHERE strRelatorioPrincipal = '" & objRelatorio.Name & "'"
               
               Set gobjBanco = New clsBanco
               
               If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
                  
                  With adoResultado
                  
                     Dim intIndex As Integer
                     Dim rptSub() As New ActiveReport
                     ReDim Preserve rptSub(1 To .RecordCount)
                     
                     intIndex = 1
                     
                     While Not .EOF
                        
                        rptSub(intIndex).LoadLayout gstrDirDocumentos & "Documentos\Relatorios\" & !strSubRelatorio & ".rpx"
                        rptSub(intIndex).AddNamedItem "clsRelatorio", New clsRelatorio
                        objActiveReports.AddNamedItem !strSubRelatorio, rptSub(intIndex)
                        .MoveNext
                        intIndex = intIndex + 1
                        
                     Wend
                  End With
                  
               End If
            
            End If
            
            'Fim da rotina criada por Tiago em 28/12/2004
            
            frmVisualizarRelatorio.WindowState = vbMaximized
            frmVisualizarRelatorio.Show
        Else
            objRelatorio.adoDataControl.Provider = ""
            objRelatorio.adoDataControl.ConnectionString = strConnectionString
            objRelatorio.adoDataControl.Source = strQuery
            Set objRelatorio.adoDataControl.Recordset = adoRelatorio
            objRelatorio.adoDataControl.ConnectionTimeout = lngIntervaloDeTempo
            objRelatorio.adoDataControl.CommandTimeout = lngIntervaloDeTempo

            
            If Trim(strTitulo) <> "" Then
                objRelatorio.Caption = strTitulo
            End If
            objRelatorio.WindowState = vbMaximized
            
            objRelatorio.Show
            
        End If
    
        
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErroImprimeRelatorio:
    ExibeDetalheErro "Erro na geração do relatório."
    Resume FimImprimeRelatorio
    
FimImprimeRelatorio:
    Screen.MousePointer = vbDefault
End Sub
    
Private Function blnSubRelatorio(strRelatorio As String) As Boolean

   Dim adoResultado As New ADODB.Recordset
   Dim strSql As String

   strSql = "SELECT * FROM tblSubRelatorios WHERE strRelatorioPrincipal = '" & strRelatorio & "'"
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
      
      If Not adoResultado.EOF Then
      
         blnSubRelatorio = True
      
      End If
      
   End If

End Function

Public Sub ImprimeRelatorioPorArray(objRelatorio As Object, _
                           Optional strArray As Variant, _
                           Optional strTitulo As String, _
                           Optional lngIntervaloDeTempo As Long, _
                           Optional xArray As XArrayDB, _
                           Optional blnxArray As Boolean = False)
    '-------------------------------------------------------------------
    ' SUB USADA PARA VISUALIZACAO DE RELATÓRIOS.
    '-------------------------------------------------------------------
    ' PARÂMETROS:
    '
    ' 1 - objRelatorio - Relatorio
    ' 2 - strArray - Array passado ao relatorio - Tipo String
    ' 4 - lngIntervaloDeTempo - Indica quanto tempo,em segundo, esperar
    '                           enquanto executa a query
    ' 5 - strArray - xArray passado ao relatorio - Tipo XArrayDb
    '-------------------------------------------------------------------

    Dim strRelatorio            As String
    Dim objActiveReports        As ActiveReport
    Dim blnExisteArquivo        As Boolean
    Dim objField As Object
    
    On Error GoTo ErroImprimeRelatorio
    
    If Not blnxArray Then
        If UBound(strArray, 2) < 0 Then Exit Sub
    Else
        If xArray.UpperBound(2) < 0 Then Exit Sub
    End If
    
    If lngIntervaloDeTempo = 0 Then
        lngIntervaloDeTempo = 30
    End If
    Screen.MousePointer = vbHourglass
    
    strRelatorio = objRelatorio.Name & ".rpx"
    
    gblnRestartRelatorio = False
        
    On Error GoTo NaoExiste
    If Dir(gstrDirDocumentos & "Documentos\Relatorios\" & strRelatorio, vbArchive) <> "" Then
        blnExisteArquivo = True
    Else
NaoExiste:
        blnExisteArquivo = False
    End If
        
    On Error GoTo ErroImprimeRelatorio

    If blnExisteArquivo Then
        Set objActiveReports = New ActiveReport
        objActiveReports.LoadLayout gstrDirDocumentos & "Documentos\Relatorios\" & strRelatorio
                
        If Not blnxArray Then
            objActiveReports.InicializaArray strArray
        Else
            objActiveReports.InicializaxArray xArray
        End If
        
        frmVisualizarRelatorio.ARViewer.ReportSource = objActiveReports
        If Trim(strTitulo) <> "" Then
            frmVisualizarRelatorio.Caption = strTitulo
        End If
            
        objActiveReports.ResetScripts
            
        objActiveReports.AddCode AdicionaCodigo
            
        objActiveReports.AddNamedItem "clsRelatorio", New clsRelatorio
            
        frmVisualizarRelatorio.WindowState = vbMaximized
        frmVisualizarRelatorio.Show
    Else
        
        If Not blnxArray Then
            objRelatorio.InicializaArray strArray
        Else
            objRelatorio.InicializaArray xArray
        End If
            
        If Trim(strTitulo) <> "" Then
            objRelatorio.Caption = strTitulo
        End If
        objRelatorio.WindowState = vbMaximized
            
        objRelatorio.Show
            
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErroImprimeRelatorio:
    ExibeDetalheErro "Erro na geração do relatório."
    Resume FimImprimeRelatorio
    
FimImprimeRelatorio:
    Screen.MousePointer = vbDefault
End Sub
    
Public Function gstrQuery(strTabela As String) As String
    Dim blnNaoOrdena    As Boolean
    Dim strCampo1       As String
    Dim strCampo2       As String
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    Dim adoCampo        As ADODB.Field
    strSql = ""
    strSql = strSql & "SELECT * FROM " & strTabela & " "
    strSql = strSql & "WHERE PKId = (SELECT MAX(PKId) FROM " & strTabela & ")"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            For Each adoCampo In adoResultado.Fields
                With adoCampo
                    If InStr(UCase(.Name), "PKID") Then
                        strCampo1 = Trim(.Name)
                    ElseIf InStr(UCase(.Name), "DESCRICAO") Or InStr(UCase(.Name), "NOME") Then
                        If Trim(strCampo2) = "" Then
                            strCampo2 = Trim(.Name)
                            Select Case .Type
                            Case adLongVarChar, adLongVarWChar
                                blnNaoOrdena = True
                            End Select
                        End If
                    End If
                End With
                If Trim(strCampo2) <> "" And Trim(strCampo1) <> "" Then
                    Exit For
                End If
            Next
            If Trim(strCampo1) = "" Then
                strCampo1 = adoResultado.Fields(0).Name
            End If
            If Trim(strCampo2) = "" Then
                strCampo2 = adoResultado.Fields(1).Name
                Select Case adoResultado.Fields(1).Type
                Case adLongVarChar, adLongVarWChar
                    blnNaoOrdena = True
                End Select
            End If
            strSql = ""
            strSql = strSql & "SELECT "
            strSql = strSql & strCampo1 & ", "
            strSql = strSql & strCampo2 & " "
            strSql = strSql & "FROM " & strTabela & " "
            If blnNaoOrdena = False Then
                strSql = strSql & "ORDER BY " & strCampo2
            End If
            gstrQuery = strSql
        End If
    End If
End Function

Public Sub GravaUsuario()
    '-------------------------------------------------------
    ' SUB USADA PARA GRAVAR A CONFIGURAÇÃO DO SISTEMA
    ' PARA O USUARIO RESPECTIVO
    '-------------------------------------------------------
    
    
    Dim strSql      As String
    On Error GoTo ErroGravaUsuario
    strSql = ""
    strSql = strSql & "UPDATE " & gstrUsuarios & " SET "
    strSql = strSql & "blnConfirmaGravacao = " & Abs(gblnConfirmaGravacao) & ", "
    strSql = strSql & "blnConfirmaExclusao = " & Abs(gblnConfirmaExclusao) & ", "
    strSql = strSql & "blnObjComGrade = " & Abs(gblnListViewComGrade) & ", "
    strSql = strSql & "blnListaAutomatica = " & Abs(gblnListagemAutomatica) & ", "
    strSql = strSql & "blnMostraDicas = " & Abs(gblnMostraDicas) & ", "
    strSql = strSql & "blnRelatorioZebrado = " & Abs(gblnRelatorioZebrado) & ", "
    strSql = strSql & "strCorZebrado = '" & gvntCorZebrado & "', "
    strSql = strSql & "bytchkFundoObjDiferente = " & Abs(gbytchkFundoObjDiferente) & ", "
    strSql = strSql & "strFundoObjInacessivel = '" & gvntFundoObjInacessivel & "', "
    strSql = strSql & "intExercicio = " & CStr(gIntExercicioUsuario) & ", "
    strSql = strSql & "dtmDtAtualizacao = " & gstrConvDtParaSql(gstrDataDoSistema(True)) & Chr(vbKeySpace)
    strSql = strSql & "WHERE PKId = " & glngCodUsr
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql
    
ErroGravaUsuario:
    If Err <> 0 Then
        On Error Resume Next
        Resume FimGravaUsuario
    End If
    
FimGravaUsuario:
End Sub

Public Function gstrSimOuNao(vntFlag As Variant) As String
    '---------------------------------------------------------------------
    ' FUNÇÃO USADA PARA RETORNAR UMA DAS PALAVRAS 'Sim' OU 'Não'
    '---------------------------------------------------------------------
    ' PARÂMETRO:
    '
    ' 1 - vntFlag
    '--------------------------------------------------------------
    If vntFlag = True Or Trim(vntFlag) = "True" Then
        gstrSimOuNao = "Sim"
    ElseIf vntFlag = False Or Trim(vntFlag) = "False" Then
        gstrSimOuNao = "Não"
    ElseIf Abs(vntFlag) = 0 Then
        gstrSimOuNao = "Não"
    ElseIf Abs(vntFlag) = 1 Then
        gstrSimOuNao = "Sim"
    End If
End Function

Public Sub LeMascacaraEspecifica()
    Dim strSql  As String
    Dim adoResultado    As ADODB.Recordset
    strSql = ""
    strSql = strSql & "SELECT strMascaraContaContabil, strMascaraCodigoOrcamentario, "
    strSql = strSql & "strMascaraElementoDespesa, bytRelatorioComEmissor,strMascaraItemDespesa "
    strSql = strSql & "FROM " & gstrConfiguracaoGeral & " "
    strSql = strSql & "WHERE PKId = 1"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                gstrMascaraContaContabil = Trim(!strMascaraContaContabil)
                gstrMascaraCodigoOrcamentario = Trim(!strMascaraCodigoOrcamentario)
                gstrMascaraElementoDespesa = Trim(!strMascaraElementoDespesa)
                gstrMascaraItemDespesa = IIf(IsNull(!strMascaraItemDespesa), "", Trim(!strMascaraItemDespesa))
                gbytRelatorioComEmissor = Abs(!bytRelatorioComEmissor)
            End If
        End With
        adoResultado.Close
        Set adoResultado = Nothing
    End If
    If Trim(gstrMascaraContaContabil) = "" Then
        gstrMascaraContaContabil = "0.0.0.0.00.0000"
    End If
    If Trim(gstrMascaraCodigoOrcamentario) = "" Then
        gstrMascaraCodigoOrcamentario = "0.0.0.0.00.0000"
    End If
    If Trim(gstrMascaraElementoDespesa) = "" Then
        gstrMascaraElementoDespesa = "0.0.0.0.00.0000"
    End If
    If Trim(gstrMascaraItemDespesa) = "" Then
        gstrMascaraItemDespesa = "0.0.0.0.00.0000"
    End If
End Sub

Public Function gvntFormatacaoEspecifica(vntValor As Variant, _
                                Optional bytObjeto As Byte) As Variant
    
    '---------------------------------------------------------------------
    ' FUNÇÃO USADA PARA FORMATAÇÃO ESPECÍFICA
    '---------------------------------------------------------------------
    ' PARÂMETRO:
    '
    ' 1 - vntValor (valor a ser formatado
    ' 2 - bytObjeto (indica se a formatação é para Conta Contábil (1),
    '                Código Orçamentário (2) ou Elemento de Despesa (3)
    '---------------------------------------------------------------------
    Dim intInd              As Integer
    Dim intLenMascara       As Integer
    Dim intLenField         As Integer
    Dim strMascara          As String
    Dim strMascaraAux       As String
    Dim STRNOME             As String
    Dim vntValorAux         As Variant
    strMascara = ""
    If Trim(gstrMascaraContaContabil) = "" Then
        LeMascacaraEspecifica
    End If
    If Trim(vntValor) <> "" Then
        Select Case UCase(TypeName(vntValor))
        Case "TEXTBOX", "COMBOBOX", "FIELD"
            STRNOME = vntValor.Name
        End Select
        If InStr(UCase(STRNOME), "CONTACONTABIL") Or bytObjeto = 1 Then
            strMascaraAux = gstrMascaraContaContabil
        ElseIf InStr(UCase(STRNOME), "CODIGOORCAMENTARIO") Or bytObjeto = 2 Then
            strMascaraAux = gstrMascaraCodigoOrcamentario
        ElseIf InStr(UCase(STRNOME), "ELEMENTODESPESA") Or bytObjeto = 3 Then
            strMascaraAux = gstrMascaraElementoDespesa
        ElseIf InStr(UCase(STRNOME), "ITEMDESPESA") Or bytObjeto = 4 Then
            strMascaraAux = gstrMascaraItemDespesa
        End If
        If Trim(strMascaraAux) = "" Then
            gvntFormatacaoEspecifica = gstrENulo(vntValor)
        Else
            For intInd = 1 To Len(Trim(strMascaraAux))
                If Mid(strMascaraAux, intInd, 1) = "0" Then
                    strMascara = strMascara & "0"
                    intLenMascara = intLenMascara + 1
                Else
                    strMascara = strMascara & "\" & Mid(strMascaraAux, intInd, 1)
                End If
                Next
            If intLenMascara > Len(gstrENulo(vntValor)) Then
                intLenField = intLenMascara - Len(gstrENulo(vntValor))
                vntValorAux = vntValor & String(intLenField, "0")
            ElseIf intLenMascara < Len(gstrENulo(vntValor)) Then
                vntValorAux = Mid(vntValor, 1, intLenMascara)
            Else
                vntValorAux = vntValor
            End If
            gvntFormatacaoEspecifica = Format(vntValorAux, strMascara)
        End If
    End If
End Function

Public Function gblnCancelaProcesso(bytFlag As Byte, _
                                    ParamArray Param()) As Boolean
'    Static blnCancelar As Boolean
'    With frmCadCheque
'        .Show
'        If bytFlag = 0 Then
'            .prgProgressao.Max = 100
'            .prgProgressao.Value = 0
'            .lblPorCento(1) = 0
'            blnCancelar = Param(0)
'            .cmd_Cancela.Enabled = True
'            If blnCancelar Then
'                .cmd_Cancela.Visible = True
''                .cmd_Cancela.SetFocus
'            Else
'                .cmd_Cancela.Visible = False
'            End If
'            DoEvents
'            If Trim(Param(1)) = "" Then
'                .fraProgressao.Caption = " Executando "
'            Else
'                .fraProgressao.Caption = Chr(vbKeySpace) & Trim(Param(1)) & Chr(vbKeySpace)
'            End If
'        ElseIf bytFlag = 1 Then
'            If Param(1) <> 0 Then
'                .prgProgressao.Value = ((Param(0) / Param(1)) * 100)
'                .lblPorCento(1) = Format(Param(0) / Param(1) * 100, "###") & "%"
'                .lblPorCento(1).Refresh
'                If blnCancelar Then
'                    DoEvents
'                End If
'                If .cmd_Cancela.Enabled = False Then
'                    Unload frmCadCheque
'                    gblnCancelaProcesso = True
'                End If
'            End If
'        Else
'            Unload frmCadCheque
'        End If
'    End With
End Function
'
Public Function gblnCarregaIcone(objObjeto As Object, _
                           ByVal strArquivo As String, _
                  Optional blnMousePointer As Boolean) As Boolean
    If InStr(strArquivo, ".") = False Then
        strArquivo = App.Path & "\" & strArquivo & ".ico"
    Else
        strArquivo = App.Path & "\" & strArquivo
    End If
    With objObjeto
        If Dir(strArquivo) <> "" Then
            If blnMousePointer Then
                .MouseIcon = LoadPicture(strArquivo)
                .MousePointer = vbCustom
            ElseIf TypeOf objObjeto Is Form Then
                .Icon = LoadPicture(strArquivo)
            Else
                .DragIcon = LoadPicture(strArquivo)
            End If
            gblnCarregaIcone = True
        End If
    End With
End Function

Public Function GetWindowsVersion() As String
    Dim strOS               As String
    Dim osvVersion          As OSVERSIONINFO
    Dim strMaintBuildInfo   As String
    osvVersion.dwOSVersionInfoSize = Len(osvVersion)
    If GetVersionEx(osvVersion) <> 0 Then
        Select Case osvVersion.dwPlatformId
            Case VER_PLATFORM_WIN32_NT
                strOS = "Windows"
                'strOS = "Windows NT"
            Case VER_PLATFORM_WIN32_WINDOWS
                strOS = "Windows"
            Case Else
                strOS = "Win32s"
        End Select
        GetWindowsVersion = strOS
    End If
End Function

Public Function GetRegString(lngKey As Long, _
                             strSubKey As String, _
                             strValueName As String) As String
    '-----------------------------------------------------------------------------
    '   FUNÇÃO UTILIZADA PARA LER A DEFINIÇÃO DE UMA CHAVE NO REGISTER DO WINDOWS
    '-----------------------------------------------------------------------------
    'PARÂMETROS:
    '   lngKey - Constante
    '   strSubKey - Caminho e nome da chave que se deseja ler
    '   strValueName - Nome do valor que se deseja ler
    '-----------------------------------------------------------------------------
    Dim strSetting  As String
    Dim lngDataLen  As Long
    Dim lngSubKey   As Long
    'Abre a chave
    If RegOpenKeyEx(lngKey, strSubKey, 0, KEY_ALL_ACCESS, lngSubKey) = ERROR_SUCCESS Then
        strSetting = Space(255)
        lngDataLen = Len(strSetting)
        'Consulta a chave em busca da definição
        If RegQueryValueEx(lngSubKey, strValueName, ByVal 0, REG_SZ, _
                           ByVal strSetting, lngDataLen) = ERROR_SUCCESS Then
            If lngDataLen > 1 Then
                GetRegString = Left(strSetting, lngDataLen - 1)
            End If
        Else
            'MsgBox "Não foi possível ler a definição da chave"
        End If
        'Fecha a chave
        RegCloseKey lngSubKey
    End If
End Function

Public Function gblnParametrosConeccaoOk(strServidor As String, _
                                         strDataBase As String, _
                                         Optional bytDatabaseType As EDatabases = SQLServer) As Boolean
    
'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Adicionado o parâmetro bytDatabaseType, valor default SQLServer (1)
'            - Adicionada leitura de chave no register afim de ler o tipo de Banco de Dados
'            - Alterada condição de OK da conexão de forma que somente seja necessária a
'              informação da variável strDataBase qdo o DB for SQL Server
' Responsável: Everton Bianchini
'******************************************************************************************

    strServidor = ""
    strDataBase = ""
    'Lê a chave no register
    
    If gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "Servidor") = "" Then
        gGravaValorRegister HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "Servidor", gstrLeValorRegister(HKEY_LOCAL_MACHINE, "SOFTWARE\CPD\AdGover\Parâmetros", "Servidor")
    End If
    
    If gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "Database") = "" Then
        gGravaValorRegister HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "Database", gstrLeValorRegister(HKEY_LOCAL_MACHINE, "SOFTWARE\CPD\AdGover\Parâmetros", "Database")
    End If
    
    If gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "DatabaseType") = "" Then
        gGravaValorRegister HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "DatabaseType", gstrLeValorRegister(HKEY_LOCAL_MACHINE, "SOFTWARE\CPD\AdGover\Parâmetros", "DatabaseType")
    End If
    
    strServidor = gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "Servidor")
    'Lê a chave no register
    strDataBase = gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "Database")
    
    'Lê a chave no register
    bytDatabaseType = Val(gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "DatabaseType"))
    
    If ((bytDatabaseType < EDatabases.SQLServer) Or (bytDatabaseType > EDatabases.Oracle)) Then
        bytDatabaseType = EDatabases.SQLServer
    End If
    
'    If Trim(strServidor) <> "" And Trim(strDataBase) <> "" Then
    If (Trim(strServidor) <> "") Then
        If (((Trim(strDataBase) <> "") And (bytDatabaseType = EDatabases.SQLServer)) Or _
            (bytDatabaseType = EDatabases.Oracle)) Then
            gblnParametrosConeccaoOk = True
        End If
    End If
End Function

Public Function gstrLeValorRegister(lngChave As Long, strSubChave As String, strValor As String) As String
    gstrLeValorRegister = ""
    If GetWindowsVersion = "Windows" Then
        'Lê a chave no register
        gstrLeValorRegister = GetRegString(lngChave, strSubChave, strValor)
    End If
End Function

Public Sub gGravaValorRegister(lngChave As Long, strSubChave As String, strValor As String, strDescricao As String)
    If Trim(strDescricao) = "" Then Exit Sub
    If GetWindowsVersion() = "Windows" Then
        'Grava a definição do database no register
        SetRegString lngChave, strSubChave, strValor, strDescricao
    End If
End Sub

Public Function gstrStringCripitografada(ByVal vntTexto As Variant, _
                                      Optional blnCripitogravar As Boolean, _
                                      Optional blnManteMinusculo As Boolean) As String
    '---------------------------------------------------------------------------------
    ' FUNÇÃO USADA PARA CRIPTOGRAFAR E/OU DESCRIPTOGRAFAR TEXTO.
    '---------------------------------------------------------------------------------
    ' PARÂMETRO:
    ' 1 - vntTexto - Texto a ser cirpitografado/descripitografado
    ' 2 - blnCripitogravar - Flag indicando se vai cripitogravar ou
    '                        descripitogravar o texto
    ' 3 - blnManteMinusculo - Flag indicando se é para manter os caracteres minúsculos
    '---------------------------------------------------------------------------------
    ' 1 - Procedimento para cripitografar:
    ' 1.1 Pega cada caracter do texto informado.
    '     For bytInd = 1 To Len(Trim(vntTexto))
    ' 1.2 Geral o código asc de cada caracter.
    '     bytAux = Asc(Mid$(vntTexto, bytInd, 1))
    ' 1.3 Formata uma string de 5 caracteres numéricos com o código asc ao quadrado.
    '     strAux = Format((bytAux ^ 2), "00000")
    ' 1.4 Pega cada caracter dessa string.
    '     For bytAux = 1 To 5
    ' 1.5 Concatena os caracteres de cada código asc da string somado 112.
    '     strTexto = strTexto & Chr(Mid(strAux, bytAux, 1) + 112)
    '--------------------------------------------------------------------------------
    Const bytCodigo112          As Byte = 112
    Dim strTexto                As String
    Dim strAux                  As String
    Dim intInd                  As Integer
    Dim bytAux                  As Byte
    Dim strAsc                  As String
    On Error Resume Next
    If blnCripitogravar Then
        For intInd = 1 To Len(Trim(vntTexto))
            bytAux = Asc(Mid$(vntTexto, intInd, 1))
            strAux = Format((bytAux ^ 2), "00000")
            bytAux = Right(Asc(Mid$(vntTexto, intInd, 1)), 1) + bytCodigo112
            strTexto = strTexto & LCase(Chr(bytAux))
            For bytAux = 1 To 5
                strTexto = strTexto & Chr(Mid(strAux, bytAux, 1) + bytCodigo112)
            Next
            bytAux = Left(Asc(Mid$(vntTexto, intInd, 1)), 1) + bytCodigo112
            strTexto = strTexto & LCase(Chr(bytAux))
        Next
    Else
        For intInd = 1 To Len(Trim(vntTexto))
            strAux = strAux & Mid(vntTexto, intInd, 1)
            If Len(strAux) Mod 7 = 0 Then
                strAsc = ""
                For bytAux = 2 To 6
                    strAsc = strAsc & Asc(Mid(strAux, bytAux, 1)) - bytCodigo112
                Next
                If blnManteMinusculo Then
                    strTexto = strTexto & Chr(Abs(Val(strAsc)) ^ 0.5)
                Else
                    strTexto = strTexto & UCase(Chr(Abs(Val(strAsc)) ^ 0.5))
                End If
                strAux = ""
            End If
        Next
    End If
    gstrStringCripitografada = strTexto
End Function

Public Sub HabilitaDesabilitaBotao(frmForm As Form, _
                                   blnFlag As Boolean, _
                        ParamArray vntListaBotao() As Variant)
                   
    '----------------------------------------------------------------------
    ' SUB USADA PARA HABILTAR OU DESABILITAR BOTOES NA BARRA DE FERRAMENTA
    '----------------------------------------------------------------------
    ' PARÂMETROS
    ' 1 - frmForm(Formulário onde está a barra de ferramenta)
    ' 2 - blnFlag(falso ou verdadeiro - habilitar ou desabilitar)
    ' 3 - vntBotao(se quiser informar qual o botão a ser tratado,
    '              caso este botão não seja informado a rotina
    '              procurará pelos botões de deletar e/ou aplicar)
    ' 4 - blnManterAplicar(flag indicando se modifica o botao aplicar)
    '----------------------------------------------------------------------
    Dim blnManterAplicar    As Boolean
    Dim blnBotaoEspecifico  As Boolean
    Dim btnBotao            As Button
    Dim vntBotaoDaLista     As Variant
    Dim objControl          As Object
    If IsMissing(vntListaBotao) = False Then
        For Each vntBotaoDaLista In vntListaBotao
            If UCase(CStr(vntBotaoDaLista)) = "FALSE" Then
                blnManterAplicar = True
            Else
                blnBotaoEspecifico = True
            End If
        Next
    End If
    For Each objControl In frmForm.Controls
        If TypeOf objControl Is Toolbar Then
            For Each btnBotao In objControl.Buttons
                If blnBotaoEspecifico Then
                    For Each vntBotaoDaLista In vntListaBotao
                        If UCase(btnBotao.Key) = UCase(vntBotaoDaLista) Then
                            btnBotao.Enabled = blnFlag
                            Exit For
                        End If
                    Next
                Else
                    Select Case UCase(btnBotao.Key)
                    Case UCase(gstrAplicar)
                        If blnManterAplicar = False Then
                            btnBotao.Enabled = blnFlag
                        End If
                    Case UCase(gstrDeletar)
                        btnBotao.Enabled = blnFlag
                    End Select
                End If
            Next
        End If
    Next
End Sub

Public Function glngNumeroDeSerie(STRNOME As String) As Long
    Dim intInd  As Integer
    Dim lngSoma As Long
    For intInd = 1 To Len(Trim(STRNOME))
        lngSoma = lngSoma + Asc(Mid(STRNOME, intInd, 1))
    Next
    glngNumeroDeSerie = lngSoma
End Function

Public Function gintIndiceCBO(cboCombo As ComboBox, _
                              vntItem As Variant, _
                     Optional blnProcList As Boolean) As Integer
    '--------------------------------------------------------------'
    ' FUNÇÃO USADA PARA PROCURAR NO ITEMDATA DO COMBOBOX O ITEM    '
    ' PASSADO COMO PARÂMETRO E RETORNAR O ÍNDICE ENCONTRADO        '
    '--------------------------------------------------------------'
    ' PARÂMETROS:                                                  '
    '                                                              '
    ' 1 - cboCombo(Combo a ser pesquisado - Tipo ComboBox)         '
    ' 2 - vntItem(Item a ser procurado - Tipo Variant)             '
    '--------------------------------------------------------------'
    Dim intInd As Integer
    gintIndiceCBO = -1
    If IsNull(vntItem) = False Then
        For intInd = 0 To cboCombo.ListCount - 1
            If blnProcList Then
                If Trim(cboCombo.list(intInd)) = Trim(vntItem) Then
                    gintIndiceCBO = intInd
                    Exit Function
                End If
            ElseIf cboCombo.ItemData(intInd) = Val(vntItem) Then
                gintIndiceCBO = intInd
                Exit Function
            End If
        Next
    End If
End Function

'AplicarGeral frmForm, objGeral, objLista, strTabela, strQueryAplicar
Public Sub AplicarGeral(frmForm As Form, _
                        objGeral As Object, _
                        lvwLista As TDBGrid, _
                        strTabela As String, _
               Optional strPKId As String, _
               Optional strQuery As String)
      
    '--------------------------------------------------------------'
    ' SUB USADA PARA AUTOMATIZAR O PREENCHIMENTO DE LISTAGEM NO    '
    ' FORMULÁRIO DE ORIGEM (CHAMADOR)
    '--------------------------------------------------------------'
    ' PARÂMETROS:                                                  '
    '                                                              '
    ' 1 - frmForm(Formulário de origem)                            '
    ' 2 - objGeral(Objeto que será preenchido)                     '
    ' 3 - lvwLista(lista que contem o dado escolhido)              '
    ' 3 - strTabela(Tabela de onde será lido os dados)             '
    ' 4 - strPKId(Chave da tabela)                                 '
    '--------------------------------------------------------------'
    
    'Dim frmFormAplicado As Form  'Formulário que está sendo atualizado
    
    On Error GoTo ErroObjeto
    
    If lvwLista.FilterActive Then Exit Sub
    
    If objGeral Is Nothing = False Then
        With frmForm
            If Trim(.txtPKId) <> "" Then
                If Trim(gstrQueryParamGeral) = "" Then
                    LeDaTabelaParaObj strTabela, objGeral, strPKId, strQuery
                Else
                    LeDaTabelaParaObj strTabela, objGeral, strPKId, gstrQueryParamGeral
                End If
                If TypeOf objGeral Is ComboBox Then
                    objGeral.ListIndex = gintIndiceCBO(objGeral, .txtPKId)
                ElseIf TypeOf objGeral Is DataCombo Then
                'reinato
                    objGeral.BoundText = Val(.txtPKId)
'                    Set frmFormAplicado = objGeral.Parent
'                    frmFormAplicado.MantemForm gstrAtualizaDataCombo, objGeral.Name
'                    Set frmFormAplicado = Nothing
                ElseIf TypeOf objGeral Is ListView Then
                    Call gblnEncontroItemNoListView(objGeral, .txtPKId, lvwTag)
                End If
                If Not TypeOf objGeral Is TrueOleDBGrid70.Column Then
                    If objGeral.Enabled Then
                        objGeral.SetFocus
                    End If
                End If
                Set objGeral = Nothing
            Else
                Exit Sub
            End If
        End With
    Else
        Exit Sub
    End If
    GoTo FimErro
    
ErroObjeto:
    Resume FimErro
    
FimErro:
    Unload frmForm
    
End Sub

Public Sub PosicionaForm(mdiPrincipal As MDIForm, _
                         frmForm As Form, _
                Optional STRTIPO As String)
    Select Case UCase(STRTIPO)
    Case "C"
        With mdiPrincipal
          frmForm.Left = Int((.Width - frmForm.Width) / 2)
          frmForm.Top = Int((.Height - frmForm.Height) / 2) - 1000
          If frmForm.Top < 0 Then
             frmForm.Top = 0
          End If
        End With
    Case Else
        If frmForm.WindowState = vbNormal Then
            frmForm.Top = 0
            frmForm.Left = 0
        End If
    End Select
End Sub

Public Function gintIndiceObjetoIndexado(objObjeto As Object, _
                                Optional blnFlgQualquerValor As Boolean) As Integer
    
    '---------------------------------------------------------------'
    ' FUNÇÃO USADA PARA RETORNAR O ÍNDICE DA OPÇÃO ESCOLHIDA NOS    '
    ' OBJETOS INDEXADOS                                             '
    '---------------------------------------------------------------'
    ' PARÂMETRO:                                                    '
    '                                                               '
    ' 1 - objCombo(Objeto indexado (array)                          '
    ' 2 - blnFlgQualquerValor(Retorna -1 para não voltar nulo       '
    '---------------------------------------------------------------'
    
    Dim bytInd  As Byte
    For bytInd = objObjeto.LBound To objObjeto.UBound
        If objObjeto(bytInd).Value = True Then
            gintIndiceObjetoIndexado = bytInd
            Exit Function
        End If
    Next
    If blnFlgQualquerValor Then
        gintIndiceObjetoIndexado = -1
    End If
End Function

Public Function gblnEncontroItemNoListView(lvwLista As ListView, _
                                           strTexto As String, _
                                  Optional intLocalProcura As Integer, _
                                  Optional intTipoProcura As Integer, _
                                  Optional intColuna As Integer = 1) As Boolean
    
    '-----------------------------------------------------------------------------
    'FUNÇÃO USADA PARA PROCURAR, NO LISTVIEW, UM TEXTO INFORMADO
    '-----------------------------------------------------------------------------
    'PARÂMETRO
    '1 - lvwLista(ListView onde será feita a procura)
    '2 - strTexto(Texto a ser procurado)
    '3 - intLocalProcura(indica se a procura será feita no item,
    '                    subitem ou tag (lvwText = 0, lvwSubitem = 1, lvwTag = 2
    '4 - intTipoProcura(indica se a procura será feita com o texto
    '                   inteiro ou parcial (lvwWhole = 0, lvwPartial 1)
    '-----------------------------------------------------------------------------
    Dim itmEncontrado    As ListItem
    Set itmEncontrado = lvwLista.FindItem(strTexto, intLocalProcura, , intTipoProcura)
    If itmEncontrado Is Nothing = False Then
        gblnEncontroItemNoListView = True
        itmEncontrado.EnsureVisible     'Scroll no ListView para selecionar item encontrado.
        itmEncontrado.Selected = True   'Seleciona o ListItem
    End If
End Function

Public Function gblnLinhaCommando(objUsuario As Object, objSenha As Object) As Boolean
    '------------------------------------------------------------
    'FUNÇÃO USADA PARA VERIFICAR SE FOI INFORMADO O USUÁRIO E A
    'SENHA NA LINHA DE COMANDO E, SE TIVER INFORMADO PREENCHER
    'OS RESPECTIVOS CAMPOS NA TELA DE LOGIN
    '------------------------------------------------------------
    'PARÂMETRO
    '1 - objUsuario - Campo do nome do usuário
    '2 - objSenha - Campo da senha do usuário
    '------------------------------------------------------------
    Dim strParam                As String
    Dim lngInd                  As Long
    Dim blnPrimeiroArgumento    As Boolean
    strParam = Trim(Command())
    objUsuario = ""
    objSenha = ""
    blnPrimeiroArgumento = True
    For lngInd = 1 To Len(strParam)
        If Mid(strParam, lngInd, 1) <> Chr(vbKeySpace) And Mid(strParam, lngInd, 1) <> vbTab Then
            If blnPrimeiroArgumento Then
                objUsuario = objUsuario & Mid(strParam, lngInd, 1)
            Else
                objSenha = objSenha & Mid(strParam, lngInd, 1)
            End If
        Else
            blnPrimeiroArgumento = False
        End If
        'retira os espaços dos campos de usuário e senha e
        'atribui o o valor verdadeiro para a função
        objUsuario = Trim(objUsuario)
        objSenha = Trim(objSenha)
        gblnLinhaCommando = True
    Next
End Function

    '--------------------------------------------------------------------------
    ' FUNÇÃO USADA PARA VERIFICAR OS FLAGS DE CONFIRMAÇÃO DE
    ' GRAVAÇÃO E EXCLUSÃO DE REGISTROS E EXIBIR UMA MENSAGEM
    ' PEDINDO CONFIRMAÇÃO CASO O FLAG SEJA VERDADEIRO.
    '--------------------------------------------------------------------------
    ' PARÂMETROS:
    '
    ' 1 - strMensagem(String contendo a mensagem, se a mensagem for específica,
    '                 ou a descrição do registro a ser gravado ou excluido
    ' 2 - blnExclusao(Se verdadeiro indica exclusão - Tipo Boolean
    ' 3 - blnInclusao(indica se mostra a plavra inclusão no lugar de gravação
    ' 4 - blnAlteracao(indica se mostra a plavra alteraçã no lugar de gravação
    '--------------------------------------------------------------------------
Public Function gblnExclusaoGravacaoOk(strModoOperacao As String, _
                              Optional strMensagem As String, _
                              Optional blnMsgEspecifica As Boolean) As Boolean
    Dim strMsg
    gblnExclusaoGravacaoOk = True
    If blnMsgEspecifica Then
        If MsgBox(strMensagem, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            gblnExclusaoGravacaoOk = False
        End If
    ElseIf gblnConfirmaExclusao And strModoOperacao = "E" Then
        strMsg = "Confirma exclusão " & Trim(strMensagem) & " ?"
        If MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            gblnExclusaoGravacaoOk = False
        End If
    ElseIf gblnConfirmaExclusao = False And strModoOperacao = "E" Then
        Exit Function
    ElseIf gblnConfirmaGravacao Then
        If strModoOperacao = "A" Then
            strMsg = "Confirma alteração " & Trim(strMensagem) & " ?"
        ElseIf strModoOperacao = "I" Then
            strMsg = "Confirma inclusão " & Trim(strMensagem) & " ?"
        Else
            strMsg = Trim(strMensagem) & " ?"
        End If
        If MsgBox(Trim(strMsg), vbYesNo + vbQuestion + vbDefaultButton1) = vbNo Then
            gblnExclusaoGravacaoOk = False
        End If
    End If
End Function

Public Function gblnHaItemMarcadoLista(lvwLista As ListView) As Boolean
    
    '---------------------------------------------------------------
    ' FUNÇÃO USADA PARA VERIFICAR SE HÁ ITEM MARCADO NO 'lvwLista'
    '---------------------------------------------------------------
    ' PARÂMETROS
    ' 1 - lvwLista(ListView em que será desmarcado)
    '---------------------------------------------------------------
    
    Dim intInd  As Integer
    With lvwLista
        For intInd = 1 To .ListItems.Count
            If .ListItems(intInd).Selected Then
                gblnHaItemMarcadoLista = True
                Exit Function
            End If
        Next
    End With
End Function

Public Function gstrItemData(objCombo As Object, _
                    Optional blnNaoRetornaZero As Boolean) As String
    '---------------------------------------------------------------'
    ' FUNÇÃO USADA PARA RETORNAR O CONTEUDO DO ITEMDATA DO COMBOBOX '
    '---------------------------------------------------------------'
    ' PARÂMETRO:                                                    '
    '                                                               '
    ' 1 - objCombo(ComboBox - Tipo ComboBox)                        '
    '---------------------------------------------------------------'
    If TypeOf objCombo Is ComboBox Then
        With objCombo
            If .ListIndex = -1 Then
                If blnNaoRetornaZero = False Then
                    gstrItemData = "0"
                Else
                    gstrItemData = "NULL"
                End If
            ElseIf .ItemData(.ListIndex) = 0 Then
                gstrItemData = .Text
            Else
                gstrItemData = .ItemData(.ListIndex)
            End If
        End With
    ElseIf TypeOf objCombo Is DataCombo Then
        With objCombo
            If Not .MatchedWithList Then
                If blnNaoRetornaZero = False Then
                    gstrItemData = "0"
                Else
                    gstrItemData = "NULL"
                End If
            Else
                gstrItemData = Val(.BoundText)
            End If
        End With
    End If
End Function

Public Sub FinalizaDragDrop(objObjeto As Object, _
                            intStatus As Integer)
    With objObjeto
        Select Case intStatus
        Case vbCancel
            TrocaInconiDoObj objObjeto, vbCancel
        Case vbBeginDrag
            TrocaInconiDoObj objObjeto, vbBeginDrag
        End Select
    End With
End Sub

Public Sub AtribuiValorDoSql(objObjeto As Object, _
                             adoCampo As ADODB.Field)

    Dim blnFlagEspecifico   As Boolean
    '-----------------------------------------------------------------'
    ' SUB USADA PARA ATRIBUIR A INFORMAÇÃO LIDA DO BANCO DE DADOS     '
    ' PARA O OBJETO RESPECTIVO NA TELA (FORM).                        '
    '-----------------------------------------------------------------'
    ' PARÂMETROS:                                                     '
    '                                                                 '
    ' 1 - objObjeto (o objeto no formulário (tela) 'ComBox, TextBox,  '
    '                OptionBoton, ChekBoton Etc).                     '
    ' 2 - adoCampo (o campo (coluna) da tabela)                       '
    '-----------------------------------------------------------------'

'******************************************************************************************
' Data: 07/03/2003
' Alteração: - Alteração na cláusula de condição para formatação dos campos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    If TypeOf objObjeto Is TextBox Or TypeOf objObjeto Is MaskEdBox Then
        If objObjeto.OLEDragMode = vbAutomatic Then
            TrocaCorObjeto objObjeto, True
        End If
    End If
    If adoCampo.Type = adBoolean Then
        If TypeOf objObjeto Is OptionButton Then
            If IsNull(adoCampo) = False Then
                VerificaIndiceObjeto objObjeto, Abs(adoCampo)
            End If
        ElseIf IsNull(adoCampo) = False Then
            objObjeto = Abs(adoCampo.Value)
        End If
    ElseIf TypeOf objObjeto Is OptionButton Then
        If IsNull(adoCampo) = False Then
            VerificaIndiceObjeto objObjeto, Val(adoCampo)
        Else
            VerificaIndiceObjeto objObjeto, 0, True
        End If
    ElseIf TypeOf objObjeto Is ComboBox Then
        objObjeto.ListIndex = gintIndiceCBO(objObjeto, adoCampo)
    ElseIf TypeOf objObjeto Is DataCombo Then
        PreencherListaDeOpcoes objObjeto, gstrVerificaCampoNulo(adoCampo)
        If objObjeto.Tag <> "" Then
            objObjeto.BoundText = gstrVerificaCampoNulo(adoCampo)
        Else
            If gstrVerificaCampoNulo(adoCampo) <> "" Then objObjeto.Text = gstrVerificaCampoNulo(adoCampo)
        End If
    ElseIf InStr("CGC,CPF", UCase(Mid(adoCampo.Name, 4, 3))) <> 0 Or _
        UCase(Mid(adoCampo.Name, 4, 4)) = "CNPJ" Then
        objObjeto = gstrCGCCPFFormatado(gstrENulo(adoCampo))
    ElseIf InStr("CEP", UCase(Mid(adoCampo.Name, 4, 3))) <> 0 Then
        objObjeto = gstrCEPFormatado(gstrENulo(adoCampo))
'    ElseIf adoCampo.Type = adVarWChar Then
    ElseIf (adoCampo.Type = adVarWChar And bytDBType = EDatabases.SQLServer) Or _
        (adoCampo.Type = adVarChar And bytDBType = EDatabases.Oracle) Then
'    Só executará a condição, qdo o DB for Oracle, se o tipo do campo for adVarChar
        objObjeto = gvntFormatacaoEspecifica(adoCampo)
    ElseIf adoCampo.Type = adCurrency Then
        objObjeto = gstrConvVrDoSql(gstrENulo(adoCampo.Value))
'    ElseIf adoCampo.Type = adNumeric Then
    ElseIf (adoCampo.Type = adNumeric) And _
        ((bytDBType = EDatabases.SQLServer) Or (bytDBType = EDatabases.Oracle And adoCampo.NumericScale > 0)) Then
'    Só executará a condição, qdo o DB for Oracle, se o campo possuir casas decimais
        objObjeto = gstrConvVrDoSql(gstrENulo(adoCampo.Value), adoCampo.NumericScale, adoCampo.Precision)
'    ElseIf adoCampo.Type = adUnsignedTinyInt Then
    ElseIf (adoCampo.Type = adUnsignedTinyInt And bytDBType = EDatabases.SQLServer) Or _
         (adoCampo.Type = adNumeric And bytDBType = EDatabases.Oracle And adoCampo.Precision <= 3) Then
'    Só executará a condição, qdo o DB for Oracle, se o campo possuir 3 ou menos dígitos
        objObjeto = Val(gstrENulo(adoCampo.Value))
    Else
        objObjeto = Replace(gstrENulo(adoCampo.Value), Chr(207), "'")
    End If
End Sub

Public Sub PreencheListaMes(ParamArray objMes())
    Dim bytInd      As Byte
    Dim vntObjeto   As Variant
    For Each vntObjeto In objMes
        Select Case UCase(TypeName(vntObjeto))
        Case "COMBOBOX"
            For bytInd = 1 To 12
                vntObjeto.AddItem gstrNomeDoMes(CStr(bytInd))
                vntObjeto.ItemData(vntObjeto.NewIndex) = bytInd
            Next
        End Select
    Next
End Sub

Public Function gstrNomeDoMes(strMes As String) As String
    '--------------------------------------------------------------'
    ' FUNÇÃO USADA PARA RETORNAR O NOME DO MES.                    '
    '--------------------------------------------------------------'
    ' PARÂMETRO:                                                   '
    '                                                              '
    ' 1 - strMes(Mes - Tipo String)                                '
    '--------------------------------------------------------------'
    Select Case Val(strMes)
        Case 1
            gstrNomeDoMes = "Janeiro"
        Case 2
            gstrNomeDoMes = "Fevereiro"
        Case 3
            gstrNomeDoMes = "Março"
        Case 4
            gstrNomeDoMes = "Abril"
        Case 5
            gstrNomeDoMes = "Maio"
        Case 6
            gstrNomeDoMes = "Junho"
        Case 7
            gstrNomeDoMes = "Julho"
        Case 8
            gstrNomeDoMes = "Agosto"
        Case 9
            gstrNomeDoMes = "Setembro"
        Case 10
            gstrNomeDoMes = "Outubro"
        Case 11
            gstrNomeDoMes = "Novembro"
        Case 12
            gstrNomeDoMes = "Dezembro"
    End Select
End Function

Public Function gstrDataPorExtenso(Optional strData As String, _
                                   Optional blnTime As Boolean, _
                                   Optional blnDiaDaSemana As Boolean) As String
    Dim strDataAux  As String
    If gblnDataValida(strData) = False Then
        strData = gstrDataDoSistema(blnTime)
    End If
    If blnDiaDaSemana Then
        strDataAux = strDataAux & gstrDiaDaSemana(strData) & ", "
    End If
    If Day(strData) = 1 Then
        strDataAux = strDataAux & "1º"
    Else
        strDataAux = strDataAux & Day(strData)
    End If
    strDataAux = strDataAux & " de " & gstrNomeDoMes(Month(strData)) & " de "
    strDataAux = strDataAux & Format(strData, "yyyy")
    If blnTime Then
        strDataAux = strDataAux & " - " & Format(strData, "hh:mm:ss")
    End If
    gstrDataPorExtenso = strDataAux
End Function

Public Sub MarcaCampo(objObjeto As Object)
    '--------------------------------------------------------------'
    ' SUB USADA MARCAR PARA SOBREPOR AS INFORMAÇÕES CONTIDAS EM    '
    ' objObjeto QUANDO DIGITAR QUALQUER CARACTER.                  '
    '--------------------------------------------------------------'
    ' PARÂMETRO:                                                   '
    '                                                              '
    ' 1 - objObjeto(Objeto a ser iniciado - Tipo Object)           '
    '--------------------------------------------------------------'
    With objObjeto
        .SelStart = 0
        If Mid(UCase(objObjeto.Name), 1, 3) = "MSK" Then
            .SelLength = Len(objObjeto.FormattedText)
        Else
            .SelLength = Len(objObjeto)
        End If
    End With
End Sub

Public Function gstrDataDoSistema(Optional blnTime As Boolean, _
                                  Optional blnNaoFormata As Boolean, _
                                  Optional blnOmiteSegundo As Boolean) As String
                                  
'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 07/03/2003
' Alteração: - Adaptação da query de retorno da data do Banco de Dados ao Oracle.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    '-------------------------------------------------
    'FUNÇÃO USADA PARA BUSCAR A DATA DO SISTEMA
    '-------------------------------------------------
    
    Dim strDataAux      As String
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    Select Case gLocalDoBancoDados
    Case bytBancoRemoto
'        strSql = "SELECT GETDATE() Data"
        strSql = "SELECT " & strGETDATE & " Data"
        If (bytDBType = EDatabases.Oracle) Then
            strSql = strSql & " FROM DUAL"
        End If
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            With adoResultado
                strDataAux = !Data
            End With
        End If
        Set adoResultado = Nothing
    End Select
    If Trim(strDataAux) = "" Then
        strDataAux = Now
    End If
    If blnNaoFormata Then
        gstrDataDoSistema = strDataAux
    Else
        gstrDataDoSistema = gstrDataFormatada(strDataAux, blnTime, blnOmiteSegundo)
    End If
End Function

Public Function gstrAnoFormatado(strAno As Variant) As String
    Dim strData As String
    If Trim(strAno) <> "" Then
        strAno = "01/01/" & strAno
        
        gstrAnoFormatado = Year(Format(strAno, "dd/mm/yyyy"))
    End If
End Function

Public Function gstrDataFormatada(vntData As Variant, _
                         Optional blnComHora As Boolean, _
                         Optional blnOmiteSegundo As Boolean, _
                         Optional blnSoHora As Boolean) As String
    '---------------------------------------------------------------
    ' FUNÇÃO USADA PARA FORMATAR DATA NO FORMATO 'DD/MM/YYYY'
    '---------------------------------------------------------------
    ' PARÂMETRO:
    ' 1 - vntData - Data a ser formatada
    ' 2 - blnComHora - Flag indicando se formata a data com hora
    ' 3 - blnOmiteSegundo - Flag indicando se omite o segundo
    ' 4 - blnSoHora - Flag para formatar apenas a data
    '--------------------------------------------------------------
    If IsNull(vntData) = False Then
        If Right(Trim(vntData), 1) = "/" Then
            If Len(Trim(vntData)) = 3 Then
                vntData = vntData & Format(gstrDataDoSistema(False, True), "mm/yy")
            ElseIf Len(Trim(vntData)) > 3 Then
                vntData = vntData & Year(gstrDataDoSistema(False, True))
            End If
        End If
        If IsDate(vntData) Then
            If blnSoHora Then
                If blnOmiteSegundo Then
                    gstrDataFormatada = Format(vntData, "hh:mm")
                Else
                    gstrDataFormatada = Format(vntData, "hh:mm:ss")
                End If
            ElseIf blnComHora Then
                If blnOmiteSegundo Then
                    gstrDataFormatada = Format(vntData, "dd/mm/yyyy hh:mm")
                Else
                    gstrDataFormatada = Format(vntData, "dd/mm/yyyy hh:mm:ss")
                End If
            Else
                gstrDataFormatada = Format(vntData, "dd/mm/yyyy")
            End If
        Else
            gstrDataFormatada = vntData
        End If
    End If
End Function

Public Sub DesmarcaIntemListView(lvwLista As ListView)
    
    '---------------------------------------------------------------
    ' SUB USADA PARA DESMARCA OS DADOS DA LISTA 'lvwLista'
    '---------------------------------------------------------------
    ' PARÂMETROS
    ' 1 - lvwLista(ListView em que será desmarcado)
    '---------------------------------------------------------------
    
    Dim intInd  As Integer
    With lvwLista
        For intInd = 1 To .ListItems.Count
            .ListItems(intInd).Selected = False
        Next
    End With
End Sub

Public Sub CaracterValido(intCaracter As Integer, _
                 Optional STRTIPO As String, _
                 Optional objObjeto As Object, _
                 Optional blnTeclaEnterIgualEnter As Boolean)
    '---------------------------------------------------------------------------
    ' SUB USADA PARA VERIFICAR A VALIDADE DE CARACTERES DIGITADOS.
    '---------------------------------------------------------------------------
    ' PARÂMETROS:
    '
    ' 1 - strTipo(Tipo do Campo - Tipo String
    '            (V)Valores, (N)Numericos), (D)Data
    ' 2 - intCaracter(Caracter digitado - Tipo Integer)
    ' 3 - objObjeto(Objeto onde está sendo digitado - Tipo Object)
    '----------------------------------------------------------------------------
    ' Obs1. Valores especiais de Caracter
    '       3 e 24 Control C e Control X
    '       22 Control V
    '       8 Back Space ou delete
    '       44 Vígula (,)
    '       45 Sinal negativo (-)
    ' Obs2. A consistencia (UCase(strTipo) = "V" And (intCaracter
    '       = 44 Or intCaracter = 45)) é para evitar a digitação de
    '       mais de uma vírgula (,) ou mais de um sinal genativo (-)
    '       em objeto de valores e a vírgula como primeiro caracter
    '       ou junto do sinal negativo
    ' Obs3. Esta função deve ser chamada do evento KeyPress
    ' Obs4. Valores para strTipo
    '       - D - Data - já forma os campos de data com '/' barra e limita o
    '             tamanho do objeto com 10 caracteres e impede a digitação de
    '             caracteres inválidos para data
    '       - E - CEP - limita o tamanho do objeto com 8 caracteres, impede a
    '             digitação de caracteres inválidos para CEP, etc.
    '       - G - Gabarito - tratamento específica para limitar o número de
    '             alternativas nos gabaritos (Concurso/Vestibular)
    '       - H - Hora - específico para objetos que receberão horas
    '       - M - converte o caracter para maiúsculo
    '       - N - Números inteiros - não permite casas decimais
    '       - U - UF - Comverte o caracter para maiúsculo e limita a tamhanho
    '             do objeto com 2 caracteres
    '       - V - Valores - permite digitação de ',' vírgula
    '----------------------------------------------------------------------------'
    Dim bytPosicaoCursor    As Byte
    On Error Resume Next
    If intCaracter = vbKeyReturn Then
        If blnTeclaEnterIgualEnter = False And gblnTeclaEnterIgualTab Then
            EnviaTeclaTab intCaracter
            Exit Sub
        End If
    End If
    Select Case intCaracter
        Case 3, 24 'Ctl C, Ctl X
            Exit Sub
        Case 22 'Ctl V
            Dim Texto As String
            Texto = Clipboard.GetText(1)
            If (STRTIPO = "V" Or STRTIPO = "N") And IsNumeric(Texto) Then
               Exit Sub
            End If
    End Select
        
    'Nino - Quado digitado o valor 0 e logo em seguida a vírgula e depois a tecla M esta função estava
    'permitindo pois M em Ascii é igual a vbKeySubtract ambos = 109
    'If UCase(strTipo) = "V" And (intCaracter = vbKeySnapshot Or intCaracter = vbKeySubtract) Then
    '44 = Vírgula
    '45 = (-)

    If UCase(STRTIPO) = "V" And (intCaracter = 44 Or intCaracter = 45) Then
        If InStr(objObjeto, Chr(intCaracter)) Then
            intCaracter = 0
        ElseIf intCaracter = 45 Then
            If Len(objObjeto) > 0 Then
                bytPosicaoCursor = objObjeto.SelStart
                objObjeto = Chr(intCaracter) & objObjeto
                objObjeto.SelStart = bytPosicaoCursor + 1
                intCaracter = 0
                Exit Sub
            End If
        ElseIf intCaracter = 44 Then
            If Trim(objObjeto) = "" Then
                objObjeto = "0"
                objObjeto.SelStart = Len(objObjeto)
            ElseIf (Mid$(objObjeto, 1, 1) = "-" And Len(objObjeto) = 1) Then
                intCaracter = 0
            End If
        End If
    End If
    Select Case UCase(STRTIPO)
    Case "V" 'para valor com casas decimais
        Select Case intCaracter
        Case vbKeyBack, vbKey0 To vbKey9, vbKeySnapshot, 45 '45 = -
            Exit Sub
        Case Else
            intCaracter = 0
            Beep
        End Select
    Case "N" 'para número inteiro
        Select Case intCaracter
        Case vbKeyBack, vbKey0 To vbKey9
            Exit Sub
        Case Else
            intCaracter = 0
            Beep
        End Select
    Case "U" 'para unidade federativa
        Select Case intCaracter
        Case vbKeyBack, vbKeyA To vbKeyZ, Asc(LCase(Chr(vbKeyA))) To Asc(LCase(Chr(vbKeyZ)))
            If Len(objObjeto) <> 2 Then
                intCaracter = Asc(UCase(Chr(intCaracter)))
            ElseIf intCaracter <> vbKeyBack Then
                intCaracter = 0
                Beep
            End If
        Case Else
            Beep
            intCaracter = 0
        End Select
    Case "D" 'para data
        VerificaCaracterParaData intCaracter, objObjeto
    Case "E" 'para CEP
        VerificaCaracterParaCEP intCaracter, objObjeto
    Case "G" 'Limite de alternativas para gabarito
        Select Case Asc(UCase(Chr(intCaracter)))
        Case vbKeyBack, vbKeyA To vbKeyF
            intCaracter = Asc(UCase(Chr(intCaracter)))
        Case Else
            intCaracter = 0
        End Select
    Case "H" 'Para hora
        VerificaCaracterParaHora intCaracter, objObjeto
    Case "M" 'Converter para maiúsculo
        intCaracter = Asc(UCase(Chr(intCaracter)))
    Case "S" 'Caracter para senha
        Select Case intCaracter
        'Aceitar apenas os caracter de números e letras
        '0 a 9 (48 a 57), "A" a "Z" (65 a 90) e "a" a "z"
        Case 1 To 7, 9 To 47, 58 To 64, 91 To 96, 123 To 255
            intCaracter = 0
        End Select
    Case "T"
        Select Case intCaracter
        Case 40, 41, 45, vbKeyBack, vbKey0 To vbKey9
            Exit Sub
        Case Else
            intCaracter = 0
            Beep
        End Select
    Case Else
        If intCaracter = vbKeyRight Then
            intCaracter = 96
        End If
    End Select
End Sub

Sub EnviaTeclaTab(intCaracter As Integer)
    If intCaracter = vbKeyReturn Then
        SendKeys "{TAB}"
        intCaracter = 0
        DoEvents
    End If
End Sub

Public Sub VerificaCaracterParaData(intCaracter As Integer, _
                                    objObjeto As Object)
    Dim intDigitado As Integer
    With objObjeto
        Select Case intCaracter
        Case vbKeyReturn
            EnviaTeclaTab intCaracter
        Case vbKeyBack, vbKeyDelete
            If .SelStart = 3 Or .SelStart = 6 Then
                SendKeys "{BACKSPACE}"
            End If
        Case 47
            If .SelStart = 2 Then
                If Mid(objObjeto, 3, 1) = "/" Then
                    intCaracter = 0
                    SendKeys "{RIGHT}"
                End If
            ElseIf .SelStart = 5 Then
                If Mid(objObjeto, 6, 1) = "/" Then
                    intCaracter = 0
                    SendKeys "{RIGHT}"
                End If
            Else
                intCaracter = 0
            End If
        Case vbKey0 To vbKey9
            If .SelStart = 0 Then
                If intCaracter > 51 Then
                    intDigitado = intCaracter
                    intCaracter = vbKey0
                    SendKeys Chr(intDigitado)
                End If
            ElseIf .SelStart = 1 Then
                If Mid(objObjeto, 1, 1) = "3" Then
                    If intCaracter > 49 Then
                        intCaracter = 0
                        Beep
                    End If
                ElseIf Mid(objObjeto, 1, 1) = "0" And intCaracter = vbKey0 Then
                    intCaracter = 0
                    Beep
                End If
            ElseIf .SelStart = 3 Then
                If intCaracter > 49 Then
                    intDigitado = intCaracter
                    intCaracter = vbKey0
                    SendKeys Chr(intDigitado)
                End If
            ElseIf .SelStart = 4 Then
                If Mid(objObjeto, 4, 1) = "1" Then
                    If intCaracter > 50 Then
                        intCaracter = 0
                        Beep
                    End If
                ElseIf Mid(objObjeto, 4, 1) = "0" And intCaracter = vbKey0 Then
                    intCaracter = 0
                    Beep
                End If
            ElseIf .SelStart = 10 Then
                intCaracter = 0
                Beep
            End If
            If .SelStart = 1 Or .SelStart = 4 Then
                SendKeys "/"
            End If
        Case Else
            intCaracter = 0
            Beep
        End Select
    End With
End Sub

Public Sub VerificaCaracterParaCEP(intCaracter As Integer, _
                                   objObjeto As Object)
    Dim intDigitado As Integer
    With objObjeto
        Select Case intCaracter
        Case vbKeyBack
            If .SelStart = 6 Then
                SendKeys "{BACKSPACE}"
            End If
        Case 45 '-'
            If .SelStart = 5 Then
                If Mid(objObjeto, 4, 1) = "-" Then
                    intCaracter = 0
                    SendKeys "{RIGHT}"
                End If
            Else
                intCaracter = 0
            End If
        Case vbKey0 To vbKey9
            If .SelStart = 9 Then
                intCaracter = 0
                Beep
            End If
            If .SelStart = 4 Then
                SendKeys "-"
            End If
        Case Else
            intCaracter = 0
            Beep
        End Select
    End With
End Sub

Public Sub VerificaCaracterParaHora(intCaracter As Integer, _
                                    objObjeto As Object)
    Dim intDigitado As Integer
    With objObjeto
        Select Case intCaracter
        Case vbKeyReturn
            EnviaTeclaTab intCaracter
        Case vbKeyBack, vbKeyDelete
            If .SelStart = 3 Then
                SendKeys "{BACKSPACE}"
            End If
        Case 58
            If .SelStart = 2 Then
                If Mid(objObjeto, 3, 1) = ":" Then
                    intCaracter = 0
                    SendKeys "{RIGHT}"
                End If
            Else
                intCaracter = 0
            End If
        Case vbKey0 To vbKey9
            If .SelStart = 0 Then
                If intCaracter > 50 Then
                    intDigitado = intCaracter
                    intCaracter = vbKey0
                    SendKeys Chr(intDigitado)
                End If
            ElseIf .SelStart = 1 Then
                If Mid(objObjeto, 1, 1) = "2" Then
                    If intCaracter > 51 Then
                        intCaracter = 0
                        Beep
                    End If
                End If
            ElseIf .SelStart = 2 Then
                If intCaracter > 52 Then
                    intCaracter = 0
                    Beep
                End If
            ElseIf .SelStart = 3 Then
                If intCaracter > 53 Then
                    intCaracter = 0
                    Beep
                End If
            ElseIf .SelStart = 5 Then
                intCaracter = 0
                Beep
            End If
            If .SelStart = 1 Then
                SendKeys ":"
            End If
        End Select
    End With
End Sub

Public Function gMontaOperacao(ByVal KeyCode As Integer) As String
Dim strModoOperacao As String

Select Case KeyCode
    Case vbKeyF2
        strModoOperacao = gstrNovo
    Case vbKeyF3
        strModoOperacao = gstrSalvar
    Case vbKeyF4
        strModoOperacao = gstrImprimir
    Case vbKeyF5
        strModoOperacao = gstrDeletar
    Case vbKeyF6
        strModoOperacao = gstrAplicar
    Case vbKeyF7
        strModoOperacao = gstrGrade
    Case vbKeyF8
        strModoOperacao = gstrRefresh
    Case vbKeyF9
        strModoOperacao = gstrFechar
End Select
gMontaOperacao = strModoOperacao
End Function

Public Function gstrCGCCPFFormatado(strCGCCPF As String, _
                           Optional strPFPJ As String) As String
    '------------------------------------------------------------------
    ' FUNÇÃO USADA PARA FORMATAR O CGC OU CPF.
    '------------------------------------------------------------------
    ' PARÂMETRO:
    '
    ' 1 - strCGCCPF(Valor digitado - Tipo String)
    ' 2 - strPFPJ(indica se pessoa jurídica -PJ- or Pessoa Física -PF-
    '------------------------------------------------------------------
    
    strCGCCPF = gstrValorSemMascara(strCGCCPF)
    If UCase(strPFPJ) = "PF" Then
        If gblnCPFOk(strCGCCPF) Then
            gstrCGCCPFFormatado = Format(strCGCCPF, "000\.000\.000\-00")
        Else
            gstrCGCCPFFormatado = strCGCCPF
        End If
    ElseIf UCase(strPFPJ) = "PJ" Then
        If gblnCGCOk(strCGCCPF) Then
            gstrCGCCPFFormatado = Format(strCGCCPF, "00\.000\.000\/0000\-00")
        Else
            gstrCGCCPFFormatado = strCGCCPF
        End If
    ElseIf gblnCGCOk(strCGCCPF) And Len(strCGCCPF) > 11 Then
        gstrCGCCPFFormatado = Format(strCGCCPF, "0#\.###\.###\/####\-##")
    
    'por Nino
    'ElseIf gblnCPFOk(strCGCCPF) Then
    '    gstrCGCCPFFormatado = Format(strCGCCPF, "###\.###\.###\-##")
    
    ElseIf gblnCPFOk(strCGCCPF) Then
        gstrCGCCPFFormatado = Format(strCGCCPF, "0##\.###\.###\-##")
    ElseIf Trim(strCGCCPF) = "0" Then
        gstrCGCCPFFormatado = ""
    Else
        gstrCGCCPFFormatado = strCGCCPF
    End If
End Function

Public Function gstrCondicaoGeral(frmForm As Form, _
                         Optional strDescricao As String) As String
    '--------------------------------------------------------------------------------
    ' FUNCAO USADA PARA PROCURAR A CHAVE PRIMÁRIA (PKId) NO FORMULÁRIO
    ' INFORMADO E MONTAR UMA STRING COM A CHAVE
    '--------------------------------------------------------------------------------
    ' PARÂMETROS:
    ' 1 - frmForm(formulário onde será procurada a chave)
    ' 2 - strDescricao(string para retornar a descrção, se houver, do registro
    '                  a ser exluido
    '--------------------------------------------------------------------------------

    Dim objControl      As Object
    If Val(frmForm.Controls("txtPKID")) <> 0 Then
        gstrCondicaoGeral = "WHERE PKId = " & Val(frmForm.Controls("txtPKID"))
    End If
    Exit Function
    For Each objControl In frmForm.Controls
        With objControl
            Select Case UCase(Trim(Mid(.Name, 4)))
            Case "PKID"
                gstrCondicaoGeral = "WHERE PKId = " & Val(objControl)
            Case "DESCRICAO", "NOME"
                strDescricao = "de: " & Trim(objControl)
            End Select
        End With
    Next
End Function

Public Function gstrCondicaoFrame(objFrame As Frame, _
                         Optional strDescricao As String) As String
    
    If Val(objFrame.Parent.Controls("txtPKID")) <> 0 Then
        gstrCondicaoFrame = "WHERE PKId = " & Val(objFrame.Tag)
    End If
    
End Function


Public Function gblnCGCOk(ByVal strNumeroCGC As String) As Boolean
    '--------------------------------------------------------------'
    ' FUNÇÃO USADA PARA VALIDAR O CGC.                             '
    '--------------------------------------------------------------'
    ' PARÂMETROS:                                                  '
    '                                                              '
    ' 1 - strNumeroCGC(Valor digitado - Tipo String)               '
    '--------------------------------------------------------------'
    Dim intSoma    As Integer
    Dim intMULT    As Integer
    Dim intDigito  As Integer
    Dim intC       As Integer
    Dim intI       As Integer
    strNumeroCGC = gstrValorSemMascara(strNumeroCGC)
    If Trim(strNumeroCGC) = "" Or Val(strNumeroCGC) = 0 Then
        Exit Function
    ElseIf Len(Trim(strNumeroCGC)) < 14 Then
        strNumeroCGC = String$(14 - Len(Trim$(strNumeroCGC)), "0") & Trim$(strNumeroCGC)
    End If
    For intC = 12 To 13
        intMULT = 2
        intSoma = 0
        For intI = intC To 1 Step -1
            intSoma = intSoma + (Val(Mid(strNumeroCGC, intI, 1) * intMULT))
            If intMULT = 9 Then
                intMULT = 2
            Else
                intMULT = intMULT + 1
            End If
        Next intI
        If intSoma Mod 11 < 2 Then
            intDigito = 0
        Else
            intDigito = 11 - (intSoma Mod 11)
        End If
        If Trim$(CStr(intDigito)) <> Mid(strNumeroCGC, (intC + 1), 1) Then
            Exit Function
        End If
    Next
    gblnCGCOk = True
End Function

Public Function gblnCPFOk(ByVal strNumeroCPF As String) As Boolean
    '--------------------------------------------------------------'
    ' FUNÇÃO USADA PARA VALIDAR O CPF.                             '
    '--------------------------------------------------------------'
    ' PARÂMETRO:                                                   '
    '                                                              '
    ' 1 - strNumeroCPF(Valor digitado - Tipo String)               '
    '--------------------------------------------------------------'
    Dim intSoma    As Integer
    Dim intMULT    As Integer
    Dim intDigito  As Integer
    Dim intC       As Integer
    Dim intI       As Integer
    strNumeroCPF = gstrValorSemMascara(strNumeroCPF)
    If Len(strNumeroCPF) = 0 Or Len(strNumeroCPF) > 11 Or Val(strNumeroCPF) = 0 Then
        Exit Function
    ElseIf Len(strNumeroCPF) < 11 Then
        strNumeroCPF = String(11 - Len(strNumeroCPF), "0") & strNumeroCPF
    End If
    For intC = 9 To 10
        intMULT = 2
        intSoma = 0
        For intI = intC To 1 Step -1
            intSoma = intSoma + (Val(Mid(strNumeroCPF, intI, 1)) * intMULT)
            intMULT = intMULT + 1
        Next intI
        If intSoma Mod 11 < 2 Then
            intDigito = 0
        Else
            intDigito = 11 - (intSoma Mod 11)
        End If
        If intDigito <> Val(Mid(strNumeroCPF, intC + 1, 1)) Then
            Exit Function
        End If
    Next
    gblnCPFOk = True
End Function

Public Function glngQtdLinhaTDBGrid(tbb_Grid As Object) As Long
    '--------------------------------------------------------------------------
    ' FUNÇÃO USADA PARA VERIFICAR E RETORNAR A QUANTIDADE DE LINHAS NO TDBGrid
    '--------------------------------------------------------------------------
    ' PARÂMETRO:
    ' 1 - tbb_Grid - Objeto onde será feita a verificação
    '--------------------------------------------------------------------------
    Dim adoResultado    As ADODB.Recordset
    Dim xadbLista       As New XArrayDB
    If tbb_Grid.DataSource Is Nothing = False Then
        If tbb_Grid.DataMode = 0 Then
            Set adoResultado = tbb_Grid.DataSource
            If adoResultado.EOF = False Then
                glngQtdLinhaTDBGrid = adoResultado.RecordCount
            End If
            Set adoResultado = Nothing
        ElseIf tbb_Grid.DataMode = 4 Then
            Set xadbLista = tbb_Grid.Array
            glngQtdLinhaTDBGrid = xadbLista.Count(1)
        End If
    End If
End Function

Public Sub ExibeMensagem(strMsg As String, _
                Optional lngIcone As Long)
    '--------------------------------------------------------------'
    ' FUNÇÃO USADA PARA EXIBIR A MENSAGEM PASSADA.                 '
    '--------------------------------------------------------------'
    ' PARÂMETRO:                                                   '
    '                                                              '
    ' 1 - strMsg(Mensagem a ser exibida - Tipo String)             '
    ' 2 - lng
    '--------------------------------------------------------------'
    Dim ScrMouse  As Integer
    ScrMouse = Screen.MousePointer
    Screen.MousePointer = vbNormal
    If lngIcone = 0 Then
        MsgBox strMsg, vbInformation
    Else
        MsgBox strMsg, lngIcone
    
    End If
    Screen.MousePointer = ScrMouse
End Sub

Public Sub LeExercicio()
    
'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 25/03/2003
' Alteração: - Substituição da variável strISNULL pela função gstrISNULL. Ver definição da
'            função para maiores esclarecimentos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim adoResultado  As ADODB.Recordset
    Dim strSql        As String
    strSql = ""
'    strSql = strSql & "SELECT ISNULL(MAX(intExercicio), DATEPART(YEAR,(GETDATE()))) intExercicio "
'    strSQL = strSQL & "SELECT " & strISNULL & "(MAX(intExercicio), " & gstrDATEPART(strYEAR, "(" & strGETDATE & ")") & ") intExercicio "
    strSql = strSql & "SELECT " & gstrISNULL("MAX(intExercicio)", gstrDATEPART(strYEAR, strGETDATE)) & " intExercicio "
    strSql = strSql & "FROM " & gstrExercicio & " "
    strSql = strSql & "WHERE bytSituacao = 1"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            gintExercicio = adoResultado!intExercicio
        End If
    End If
End Sub

Public Function gblnBancoDadosOK(strServidor As String, _
                                 strDataBase As String, _
                                 strUser As String, _
                                 strPassword As String, _
                        Optional blnErroServidor As Boolean, _
                        Optional blnErroDataBase As Boolean) As Boolean
                        
'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Adicionada verificação a fim de determinar qual a string de conexão a ser
'            construída.
'            - Chamada ao procedimento de inicialização das variáveis de comandos nativos.
'            - Alteração do tratamento de mensagens de erros do Banco de Dados.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strConeccao As String
   
    On Error GoTo ErroAbreBancoDados
    bytLocalBD = 1
    gLocalDoBancoDados = bytLocalBD
    
    Set gcncADOMain = New ADODB.Connection
    
    Select Case gLocalDoBancoDados
    Case bytBancoLocal
        strConeccao = "DRIVER={Microsoft Access Driver (*.mdb)};" & _
                      "DBQ=" & strDataBase & ";" & _
                      "DefaultDir=" & strServidor & ";" & "UID=admin;PWD=;"
    Case bytBancoRemoto
'       strConeccao = "Driver={SQL Server};Server=" & strServidor & ";" & _
'                     "DataBase=" & strDataBase & ";UID=" & strUser & ";PWD=" & strPassword & ";"
        
        If (bytDBType = EDatabases.SQLServer) Then
            'Para sql 7
            strConeccao = "Provider=SQLOLEDB.1;Current Language=English;Persist Security Info=True;User ID=" & strUser & ";PWD=" & strPassword & ";Initial Catalog=" & strDataBase & ";Data Source=" & strServidor
            
            'Para sql 6.5
            'strConeccao = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=" & strUser & ";PWD=" & strPassword & ";Initial Catalog=" & strDataBase & ";Data Source=" & strServidor
        
        ElseIf (bytDBType = EDatabases.Oracle) Then
            strConeccao = "Provider=MSDAORA.1;Password=" & strPassword & ";User ID=" & strUser & ";Data Source=" & strServidor & ";Persist Security Info=True"
            
        End If
              
        gcncADOMain.ConnectionTimeout = 300
        gcncADOMain.CommandTimeout = 300
        
    End Select
    
    gcncADOMain.ConnectionString = strConeccao
    gcncADOMain.IsolationLevel = adXactReadCommitted
    gcncADOMain.Open
    
    Call IniciaVarsCmdsNativosDB
    
    gblnBancoDadosOK = True
    Exit Function
    
ErroAbreBancoDados:
    If InStr(Err.Description, "Cannot open database") Then
        ExibeDetalheErro "SQL não pode abrir ou não encontrou o banco de dados '" & strDataBase & "'"
        blnErroDataBase = True
'    ElseIf InStr(Err.Description, "SQL server not found") Then
    ElseIf InStr(Err.Description, "SQL server not found") Or _
        InStr(Err.Description, "could not resolve service name") Then
        ExibeDetalheErro "SQL não encontrou o servidor '" & strServidor & "'"
        blnErroDataBase = True
    ElseIf InStr(Err.Description, "SQL Server does not exist or access denied") Then
        ExibeDetalheErro "Servidor de SQL '" & Trim(strServidor) & "' não existe ou está inacessível"
        blnErroServidor = True
'    ElseIf InStr(Err.Description, "Login failed for user") Then
    ElseIf InStr(Err.Description, "Login failed for user") Or _
        InStr(Err.Description, "logon denied") Then
        ExibeDetalheErro "Usuário e/ou senha inválidos." & Chr(13) & Chr(13) & "Tente novamente."
    Else
        ExibeDetalheErro "Erro desconhecido na conexão com o banco"
        blnErroServidor = True
    End If
End Function


Public Function gblnDataValida(vntData As Variant, _
                      Optional blnExibeMensagem As Boolean, _
                      Optional blnNaoPodeSerSuperior As Boolean) As Boolean
    '---------------------------------------------------------------------------
    ' FUNÇÃO USADA PARA VERIFICAR A VALIDADE DE DATA
    '---------------------------------------------------------------------------
    ' PARÂMETROS:
    ' 1 - vntData - data a ser testada
    ' 2 - blnExibeMensagem - flag indicando se exibe mensagem de data inválida
    '---------------------------------------------------------------------------
    ' OBS: Data para o SQL não pode ser inferior a 01/01/1753
    '---------------------------------------------------------------------------
    If IsDate(vntData) Then
        If blnNaoPodeSerSuperior And (CVDate(vntData) > CVDate(gstrDataDoSistema)) Then
            If blnExibeMensagem Then
                ExibeMensagem "A data não pode ser superior a data de hoje"
            End If
        ElseIf CVDate(vntData) < CVDate("01/01/1753") Then
            If blnExibeMensagem Then
                ExibeMensagem "A data não pode ser inferior a 1º de janeiro de 1753"
            End If
        Else
            gblnDataValida = True
       End If
    ElseIf blnExibeMensagem Then
        If Trim(vntData) <> "" Then
            ExibeMensagem Trim(vntData) & " não é uma data correta"
        Else
            ExibeMensagem "Data não pode ficar vazia"
        End If
    End If
End Function


Public Function gstrConvVrParaSql(vntNumero As Variant) As String

    '----------------------------------------------------------
    ' FUNÇÃO USADA PARA TROCAR A VIRGULA POR PONTO EM VALORES
    ' COM CASA DECIMAIS DECIMAIS.
    '----------------------------------------------------------
    ' 1 - vntNumero(Número a ser formatado
    '----------------------------------------------------------

    Dim intInd  As Integer
    Dim strAux  As String
    Dim vntAux  As Variant
    vntAux = gstrENulo(vntNumero)
    For intInd = 1 To Len(vntAux)
        If Mid(vntAux, intInd, 1) = "," Then
           strAux = strAux & "."
        ElseIf Mid(vntAux, intInd, 1) <> "." Then
           strAux = strAux & Mid(vntAux, intInd, 1)
        End If
    Next
    If Trim(strAux) = "" Then
        gstrConvVrParaSql = "Null"
    Else
        gstrConvVrParaSql = strAux
    End If
End Function

Public Function glngPegaUltimaChave(strTabela As String, _
                                     strCampo As String, _
                            Optional strCampoCond1 As String, _
                            Optional vntValor1 As Variant, _
                            Optional strCampoCond2 As String, _
                            Optional vntValor2 As Variant) As Long
    '--------------------------------------------------------------'
    ' FUNÇÃO USADA PARA RETORNAR A ÚLTIMA CHAVE DA TABELA.         '
    '--------------------------------------------------------------'
    ' PARÂMETROS:                                                  '
    '                                                              '
    ' 1 - sTAbela(Tabela - Tipo String)                            '
    ' 2 - sCampo(Campo chave da tabela - Tipo String)              '
    '--------------------------------------------------------------'
    
'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 25/03/2003
' Alteração: - Substituição da variável strISNULL pela função gstrISNULL. Ver definição da
'            função para maiores esclarecimentos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim adoResultado  As ADODB.Recordset
    Dim strSql        As String
    Dim mobjAuxBanco  As Object

    strSql = ""
'    strSql = strSql & "SELECT ISNULL(MAX(" & Trim(strCampo) & "), 0) Maximo "
'    strSQL = strSQL & "SELECT " & strISNULL & "(MAX(" & Trim(strCampo) & "), 0) Maximo "
    strSql = strSql & "SELECT " & gstrISNULL("MAX(" & Trim(strCampo) & ")", "0") & " Maximo "
    strSql = strSql & "FROM " & strTabela
    If Trim(strCampoCond1) <> "" Then
        strSql = strSql & " WHERE " & strCampoCond1 & " = " & vntValor1
    End If
    If Trim(strCampoCond2) <> "" Then
        strSql = strSql & " AND " & strCampoCond2 & " = " & vntValor2
    End If
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        glngPegaUltimaChave = adoResultado!Maximo
        Set gobjBanco = Nothing
        adoResultado.Close
        Set adoResultado = Nothing
    End If
End Function

Public Function gstrValorSemMascara(vntTexto As Variant) As String
    '--------------------------------------------------------------'
    ' FUNÇÃO USADA PARA RETIRAR OS CARACTERES DE FORMATAÇÃO        '
    '--------------------------------------------------------------'
    ' PARÂMETRO:                                                   '
    '                                                              '
    ' 1 - vntTexto(Texto - Tipo Variant)                           '
    '--------------------------------------------------------------'
    Dim strAux  As String
    Dim intInd  As Integer
    'If IsNull(vntTexto) = False Then
    If Trim(gstrENulo(vntTexto)) <> "" Then
        Do
            intInd = intInd + 1
            If InStr("0123456789", Mid(vntTexto, intInd, 1)) Then
                strAux = strAux & Mid(vntTexto, intInd, 1)
            End If
        Loop Until intInd > Len(vntTexto)
    End If
    gstrValorSemMascara = strAux
End Function

Public Function gstrVerificaCampoNulo(vntCampo As Variant, _
                             Optional blnText As Boolean) As String
    '---------------------------------------------------------------'
    ' FUNÇÃO USADA PARA VERIFICAR SE O CAMPO ESTA NULL.             '
    '---------------------------------------------------------------'
    ' PARÂMETRO:                                                    '
    '                                                               '
    ' 1 - vntCampo(Campo - Tipo Variant)                            '
    '---------------------------------------------------------------'
    If blnText Then
        If vntCampo.ColumnSize = 0 Then
            gstrVerificaCampoNulo = ""
        Else
            gstrVerificaCampoNulo = Trim(vntCampo)
        End If
    ElseIf IsNull(vntCampo) Then
        gstrVerificaCampoNulo = ""
    Else
        gstrVerificaCampoNulo = Trim(vntCampo)
    End If
End Function

Public Function gstrENulo(vntCampo As Variant, _
                 Optional blnText As Boolean, Optional blnRetornaStringNULL As Boolean) As String
    '---------------------------------------------------------------'
    ' FUNÇÃO 'é nulo' USADA PARA VERIFICAR SE O CAMPO ESTA NULL.             '
    '---------------------------------------------------------------'
    ' PARÂMETRO:                                                    '
    '                                                               '
    ' 1 - vntCampo(Campo - Tipo Variant)                            '
    '---------------------------------------------------------------'
    On Error GoTo ErroENulo
    If blnText Then
        If vntCampo.ColumnSize = 0 Then
            gstrENulo = ""
        Else
            gstrENulo = Trim(vntCampo)
        End If
    ElseIf IsNull(vntCampo) Then
        gstrENulo = ""
    Else
        gstrENulo = Trim(vntCampo)
    End If
    
    If blnRetornaStringNULL And (vntCampo = "" Or IsNull(vntCampo)) Then
        gstrENulo = "NULL"
    End If
    
    Exit Function
ErroENulo:
    Resume FimENulo
    
FimENulo:

End Function

Public Function gstrConvVrDoSql(vntNumero As Variant, _
                     Optional vntNumCasaDecimal As Variant, _
                     Optional intNumCasaInteira As Integer, _
                     Optional blnRetornaZero As Boolean) As String
    '---------------------------------------------------------------
    ' FUNÇÃO USADA PARA TROCAR OS PONTOS (.) POR VÍRGULA E FORMATAR
    ' O NÚMERO INFORMADO CONFORME A MÁSCARA INDICADA
    '---------------------------------------------------------------
    ' PARÂMETRO:
    '
    ' 1 - vntNumero(Número a ser formatado)
    ' 2 - Parametro(Pode conter até 3 valor que são:
    '     a) número de casas decimais;
    '     b) número de dígitos (parte inteira do número)
    '     c) flag indicando se a função retorna 0 (zero) ou nulo
    '---------------------------------------------------------------
                     
    Dim intInd              As Integer
    Dim bytAux              As Byte
    Dim blnJaAchouPonto     As Boolean
    Dim strAux              As String
    Dim strMascara          As String
    'Número de casas decimais
    If IsNumeric(vntNumCasaDecimal) Then
        If vntNumCasaDecimal > 0 Then
            strMascara = "." & String(vntNumCasaDecimal, "0")
        End If
        If intNumCasaInteira = 0 Then
            intNumCasaInteira = 12
        End If
    End If
    'Número de dígitos (parte inteira)
    If Val(intNumCasaInteira) > 0 Then
        For bytAux = 1 To intNumCasaInteira
            If (bytAux - 1) Mod 3 = 0 And (bytAux - 1) > 0 Then
                strMascara = "," & strMascara
            ElseIf intInd = 1 Then
                strMascara = "0" & strMascara
            ElseIf Left(strMascara, 1) = "." Then
                strMascara = "0" & strMascara
            Else
                strMascara = "#" & strMascara
            End If
        Next
    End If
    'Verifica se foi passado valor
    If Len(Trim(vntNumero)) = 0 Then
        If blnRetornaZero Then
            vntNumero = 0
        Else
            Exit Function
        End If
    End If
    'Verifica a máscara de formatação
    If Trim(strMascara) = "" Then
        strMascara = "###,###,###,##0.00"
    ElseIf Left(strMascara, 1) = "," Then
        strMascara = Mid(strMascara, 2)
    End If
    If IsNull(vntNumero) Then
        vntNumero = ""
    End If
    'Troca os pontos (.) por vírgula (,)
    For intInd = Len(Trim(vntNumero)) To 1 Step -1
        If Mid(vntNumero, intInd, 1) = "." Or Mid(vntNumero, intInd, 1) = "," Then
            If blnJaAchouPonto = False Then
                strAux = "," & strAux
                blnJaAchouPonto = True
            End If
        Else
            strAux = Mid(vntNumero, intInd, 1) & strAux
        End If
    Next
    If Val(gstrConvVrParaSql(strAux)) > 9999999999.99 Then
        strAux = Mid(strAux, 1, 12)
    End If
    'Formata o número conforme a máscara informada
    'gstrConvVrDoSql = Format(strAux, strMascara) Mudado por que quando era 0 retornava  "vazio"
    gstrConvVrDoSql = IIf(Format(strAux, strMascara) <> "", Format(strAux, strMascara), "0")
End Function

Public Function gvntConvVrDoSql(vntNumero As Variant, _
                     ParamArray Parametro()) As Variant
                     
    '---------------------------------------------------------------
    ' FUNÇÃO USADA PARA TROCAR OS PONTOS (.) POR VÍRGULA E FORMATAR
    ' O NÚMERO INFORMADO CONFORME A MÁSCARA INDICADA
    '---------------------------------------------------------------
    ' PARÂMETRO:
    '
    ' 1 - vntNumero(Número a ser formatado)
    ' 2 - Parametro(Pode conter até 3 valor que são:
    '     a) número de casas decimais;
    '     b) número de dígitos (parte inteira do número)
    '     c) flag indicando se a função retorna 0 (zero) ou nulo
    '---------------------------------------------------------------
                     
    Dim intInd              As Integer
    Dim bytAux              As Byte
    Dim blnRetornaZero      As Boolean
    Dim blnJaAchouPonto As Boolean
    Dim vntItem             As Variant
    Dim strAux              As String
    Dim strMascara          As String
    For Each vntItem In Parametro
        intInd = intInd + 1
        Select Case intInd
        Case 1
    'Número de casas decimais
            If vntItem > 0 Then
                strMascara = "." & String(vntItem, "0")
            End If
        Case 2
    'Número de dígitos (parte inteira)
            If vntItem > 0 Then 'Segundo parâmetro
                For bytAux = 1 To vntItem
                    If (bytAux - 1) Mod 3 = 0 And (bytAux - 1) > 0 Then
                        strMascara = "," & strMascara
                    ElseIf intInd = 1 Then
                        strMascara = "0" & strMascara
                    ElseIf Left(strMascara, 1) = "." Then
                        strMascara = "0" & strMascara
                    Else
                        strMascara = "#" & strMascara
                    End If
                Next
            ElseIf Trim(strMascara) <> "" Then
                strMascara = "0" & strMascara
            End If
    'Indica se retorna o valor zero ou nulo
        Case 3
            blnRetornaZero = vntItem 'Terceiro parâmetro
        End Select
    Next vntItem
    'Verifica se foi passado valor
    If Len(Trim(vntNumero)) = 0 Then
        If blnRetornaZero Then
            vntNumero = 0
        Else
            Exit Function
        End If
    End If
    'Verifica a máscara de formatação
    If Trim(strMascara) = "" Then
        strMascara = "###,###,###,##0.00"
    ElseIf Left(strMascara, 1) = "," Then
        strMascara = Mid(strMascara, 2)
    End If
    If IsNull(vntNumero) Then
        vntNumero = ""
    End If
    'Troca os pontos (.) por vírgula (,)
    For intInd = Len(Trim(vntNumero)) To 1 Step -1
        If Mid(vntNumero, intInd, 1) = "." Or Mid(vntNumero, intInd, 1) = "," Then
            If blnJaAchouPonto = False Then
                strAux = "," & strAux
                blnJaAchouPonto = True
            End If
        Else
            strAux = Mid(vntNumero, intInd, 1) & strAux
        End If
    Next
    'Formata o número conforme a máscara informada
    gvntConvVrDoSql = Format(strAux, strMascara)
End Function

Public Sub VerificaIndiceObjeto(objObjeto As Object, _
                                bytValor As Byte, _
                                Optional blnNulo As Boolean = False)
                                
    '---------------------------------------------------------------
    ' SUB USADA PARA COMPARAR O INFORMADO COM ÍNDICE DO OBJETO E
    ' FAZER VERDADEIRO O REPECTIVO OBJETO
    '---------------------------------------------------------------
    ' PARÂMETRO:
    '
    ' 1 - objObjeto
    ' 2 - bytValor
    '---------------------------------------------------------------
                                
    With objObjeto
        If blnNulo Then
            objObjeto = False
        Else
            If .Index = bytValor Then
                objObjeto = True
            End If
        End If
    End With
End Sub

Public Sub SetRegString(hKey As Long, strSubKey As String, strValueName As String, strSetting As String)
    '-----------------------------------------------------------------------------
    '   FUNÇÃO UTILIZADA PARA GRAVAR UMA CHAVE E SUA DEFINIÇÃO NO REGISTER DO WINDOWS
    '-----------------------------------------------------------------------------
    'PARÂMETROS:
    '   HKEY - Constante
    '   strSubKey - Caminho e nome da chave que se deseja criar
    '   strValueName - Nome do valor
    '   strSetting - Valor que que se deseja gravar no parâmetro acima
    '-----------------------------------------------------------------------------
    Dim hNewHandle As Long
    Dim lpdwDisposition As Long
    Dim SAttributes As SECURITY_ATTRIBUTES
    
    SAttributes.bInheritHandle = 0
    SAttributes.lpSecurityDescriptor = 0
    SAttributes.nLength = 0
    
    'Cria e abre a chave. Se OK , grava  os dados
    If RegCreateKeyEx(hKey, strSubKey, 0, strValueName, 0, KEY_ALL_ACCESS, SAttributes, hNewHandle, lpdwDisposition) = ERROR_SUCCESS Then
        If RegSetValueEx(hNewHandle, strValueName, 0, REG_SZ, ByVal strSetting, Len(strSetting)) <> ERROR_SUCCESS Then
            MsgBox "Erro ao gravar a definição da chave."
        End If
    Else
        MsgBox "Erro ao cria chave no register do windows"
    End If
    'Fecha a chave
    RegCloseKey hNewHandle
End Sub

Public Sub TrocaCorObjeto(objControle As Object, _
                 Optional blnAlterando As Boolean, _
                 Optional blnLimpaObjeto As Boolean)
'******************************************************************************************
' Data: 08/09/2003
' Alteração: - Implementado na função a funcionalidade de desabilitar os objetos dentro da
'            frame. Assim conseguiremos centralizar as mudanças de estado dos controles
' Responsável: Gustavo Monteiro
'------------------------------------------------------------------------------------------
    
    
    'Declarado para verificar os objetos filhos dentro do frame
    Dim objFilho As Object
    Dim objContainer As Object
    If TypeOf objControle Is OptionButton Then
        objControle.Enabled = Not blnAlterando
        Exit Sub
    End If
    
    'Verifica se o objeto é um frame. Caso verdadeiro, varre os filhos para (des)habilita-los
    If TypeOf objControle Is Frame Then
            
        For Each objFilho In objControle.Parent.Controls
            
            Set objContainer = objFilho.Container
                
                If Not Not (objContainer.Name) = objControle.Name And objFilho.Name <> objControle.Name _
                    And Not TypeOf objFilho Is Label And Not TypeOf objFilho Is OptionButton And Not TypeOf objFilho Is Image Then
                If gbytchkFundoObjDiferente And blnAlterando Then
                    objFilho.BackColor = Val(gvntFundoObjInacessivel)
                Else
                    objFilho.BackColor = vbWindowBackground
                End If
                    
                objFilho.Enabled = Not blnAlterando
            End If
            
            If blnLimpaObjeto Then
                If objContainer.Name = objControle.Name And objFilho.Name <> objControle.Name And Not TypeOf objFilho Is Label Then
                    With objFilho
                        If TypeOf objFilho Is TextBox Or _
                           TypeOf objFilho Is ComboBox Or _
                           TypeOf objFilho Is DataCombo Then
                            If objFilho.OLEDragMode = vbAutomatic Then
                                TrocaCorObjeto objFilho
                            End If
                        End If
                        If TypeOf objFilho Is Label Then
                            If objFilho.Tag = 1 Then
                                TrocaCorObjeto objFilho
                                objFilho = ""
                            End If
                        Else 'If Trim(Mid(.Name, 4, 1)) <> "_" Then
                            If TypeOf objFilho Is TextBox Then
                                If objFilho.CausesValidation Then
                                    objFilho.Text = ""
                                End If
                            ElseIf TypeOf objFilho Is MaskEdBox Then
                                Dim strMask As String
                                strMask = objFilho.Mask
                                objFilho.Mask = ""
                                objFilho.Text = ""
                                objFilho.Mask = strMask
                            ElseIf TypeOf objFilho Is ComboBox Then
                                If objFilho.IntegralHeight Then
                                    objFilho.ListIndex = -1
                                End If
                            ElseIf TypeOf objFilho Is DataCombo Then
                                If objFilho.IntegralHeight Then
                                    If objFilho.Style = dbcDropdownList Then
                                        objFilho.BoundText = ""
                                    Else
                                        objFilho.Text = ""
                                    End If
                                End If
                            ElseIf TypeOf objFilho Is CheckBox Then
                                objFilho.Value = 0
                            ElseIf TypeOf objFilho Is OptionButton Then
                                If objFilho.CausesValidation Then
                                    objFilho.Value = True
                                Else
                                    objFilho.Value = False
                                End If
                            End If
                        End If
                    End With
                End If
            End If
            Set objContainer = Nothing
                
        Next
        
        objControle.Enabled = Not blnAlterando
        
        Exit Sub
    End If
    
    If blnAlterando Then
        
        
        If UCase(Mid(objControle.Name, 1, 4)) <> "CMD_" Then
            If gbytchkFundoObjDiferente Then
                objControle.BackColor = Val(gvntFundoObjInacessivel)
            Else
                objControle.BackColor = vbWindowBackground
            End If
        End If
        objControle.Enabled = False
    ElseIf objControle.OLEDropMode = 0 Then
        If UCase(Mid(objControle.Name, 1, 4)) <> "CMD_" Then
            objControle.BackColor = vbWindowBackground
        End If
        objControle.Enabled = True
    ElseIf gbytchkFundoObjDiferente Then
        objControle.BackColor = Val(gvntFundoObjInacessivel)
    ElseIf UCase(Mid(objControle.Name, 1, 4)) <> "CMD_" Then
        objControle.BackColor = vbWindowBackground
    End If
End Sub

Public Sub AtivaPastaDeObjeto(tabPasta1 As SSTab, _
                              bytTabAtivo1 As Byte, _
                     Optional tabPasta2 As SSTab, _
                     Optional bytTabAtivo2 As Byte)
    With tabPasta1
        If .Tab <> bytTabAtivo1 Then
            If .TabEnabled(bytTabAtivo1) Then
                .Tab = bytTabAtivo1
            End If
        End If
    End With
    If tabPasta2 Is Nothing = False Then
        With tabPasta2
            If .Tab <> bytTabAtivo2 Then
                If .TabEnabled(bytTabAtivo2) Then
                    .Tab = bytTabAtivo2
                End If
            End If
        End With
    End If
End Sub

Public Sub LimpaObjeto(frmForm As Form, _
              Optional blnAlterando As Boolean, _
              Optional blnManterAplicar As Boolean, _
              Optional blnNaoFocarObj As Boolean)
                  
    '-------------------------------------------------------------------------------
    ' FUNÇÃO USADA PARA LIMPAR OS OBJETOS NOS FORMULÁRIOS
    '-------------------------------------------------------------------------------
    ' PARÂMETRO:
    '
    ' 1 - frmForm(formulário onde será limpado os objetos)
    ' 2 - blnAlterando(flag que indica alteração ou inclusão de registro,
    '                  retornar falso indicando que os objetos estão limpos e
    '                  um novo registro sera incluido)
    ' 3 - blnManterAplicar(sem utilização)
    '------------------------------------------------------------------------------
              
    Dim objControl      As Object
    Dim strMascara      As String
    'On Error Resume Next
    For Each objControl In frmForm.Controls
        With objControl
            If TypeOf objControl Is TextBox Or _
               TypeOf objControl Is ComboBox Or _
               TypeOf objControl Is DataCombo Then
                If objControl.OLEDragMode = vbAutomatic Then
                    TrocaCorObjeto objControl
                End If
            End If
            If TypeOf objControl Is Label Then
                If objControl.Tag = 1 Then
                    TrocaCorObjeto objControl
                    objControl = ""
                End If
            ElseIf Trim(Mid(.Name, 4, 1)) <> "_" Then
                If TypeOf objControl Is TextBox Then
                    If objControl.CausesValidation Then
                        objControl.Text = ""
                    End If
                ElseIf TypeOf objControl Is MaskEdBox Then
                    Dim strMask As String
                    strMask = objControl.Mask
                    objControl.Mask = ""
                    objControl.Text = ""
                    objControl.Mask = strMask
                ElseIf TypeOf objControl Is ComboBox Then
                    If objControl.IntegralHeight Then
                        objControl.ListIndex = -1
                    End If
                ElseIf TypeOf objControl Is DataCombo Then
                    If objControl.IntegralHeight Then
                        If objControl.Style = dbcDropdownList Then
                            objControl.BoundText = ""
                        Else
                            objControl.Text = ""
                        End If
                    End If
                ElseIf TypeOf objControl Is CheckBox Then
                    'Nino - Condição Causes Validation para que se possa escolher se limpa a check ou não
                    If objControl.CausesValidation Then
                        objControl.Value = 0
                    End If
                ElseIf TypeOf objControl Is OptionButton Then
                    If objControl.CausesValidation Then
                        objControl.Value = True
                    Else
                        objControl.Value = False
                    End If
                ElseIf TypeOf objControl Is ListBox Then
                    objControl.Clear
                End If
                
                If blnNaoFocarObj = False Then
                    If TypeOf objControl Is Image = False And TypeOf objControl Is ImageList = False Then
                        If objControl.TabIndex = 0 Then
                            If objControl.Visible = True Then
                                If objControl.Enabled = True Then
                                    objControl.SetFocus
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
    blnAlterando = False
End Sub

Public Sub MostraCaixaCores(Optional objControl As Object, _
                            Optional blnCorDeFundo As Boolean, _
                            Optional vntRetorno As Variant)
    On Error GoTo ErroCaixaCores
    With MDIMenu.dlgConfigura
        .flags = cdlCCRGBInit
        .CancelError = True
        .ShowColor
        If IsMissing(vntRetorno) = False Then
            vntRetorno = .Color
        ElseIf blnCorDeFundo Then
            objControl.BackColor = .Color
        Else
            objControl.ForeColor = .Color
        End If
    End With
    
ErroCaixaCores:
    Resume FimCaixaCores
FimCaixaCores:

End Sub

Sub CarregaForm(frmForm As Form, _
                Optional objObjeto As Object, _
                Optional strQuery As String, Optional intMode As Integer)
    
    Screen.MousePointer = 11
    gstrQueryParamGeral = strQuery
    Set gobjGeral = objObjeto
    If frmForm.WindowState <> vbMaximized Then
        frmForm.Left = 0
        frmForm.Top = 0
        frmForm.BorderStyle = vbFixedDialog
    End If
    TrocaInconiDoObj frmForm, 3
    
    If intMode = 1 Then Screen.MousePointer = 0
    
    frmForm.Show intMode
    If intMode = 0 Then
        frmForm.SetFocus
    End If
    Screen.MousePointer = 0
End Sub

Public Function gstrLeArquivoLogin(strChavePrincipal As String, _
                                   strChaveSecundaria As String, _
                                   strDefault As String) As String
    Dim strCaminhoArquivo   As String
    Dim strCoord            As String
    On Error GoTo Err_LeArquivoLogin
    strCaminhoArquivo = App.Path & "\Login.ini"
    strCoord = Space$(255)
    Call GetPrivateProfileString(strChavePrincipal, strChaveSecundaria, _
                                 strDefault, strCoord, Len(strCoord), _
                                 strCaminhoArquivo)
    gstrLeArquivoLogin = Trim(strCoord)
    Exit Function
    
Err_LeArquivoLogin:
    gstrLeArquivoLogin = ""
    
End Function

Public Function gstrGravaArquivoLogin(strChavePrincipal As String, _
                                      strChaveSecundaria As String, _
                                      strValor As String)
    Dim strCaminhoArquivo   As String
    On Error GoTo GravaArquivoLogin
    strCaminhoArquivo = App.Path & "\Login.ini"
    Call WritePrivateProfileString(strChavePrincipal, strChaveSecundaria, _
                                   Trim$(strValor), strCaminhoArquivo)
GravaArquivoLogin:

End Function

Public Function ShellEx(ByVal sFile As String, _
               Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, _
               Optional ByVal sParameters As String = "", _
               Optional ByVal sDefaultDir As String = "", _
               Optional sOperation As String = "open", _
               Optional Owner As Long = 0) As Boolean
    Dim lR As Long
    Dim lErr As Long, sErr As Long
    If (InStr(UCase$(sFile), ".EXE") <> 0) Then
        eShowCmd = 0
    End If
    On Error Resume Next
    If (sParameters = "") And (sDefaultDir = "") Then
        lR = ShellExecuteForExplore(Owner, sOperation, sFile, 0, 0, essSW_SHOWNORMAL)
    Else
        lR = ShellExecute(Owner, sOperation, sFile, sParameters, sDefaultDir, eShowCmd)
    End If
    If (lR < 0) Or (lR > 32) Then
        ShellEx = True
    Else
        ' raise an appropriate error:
        lErr = vbObjectError + 1048 + lR
        Select Case lR
        Case 0
            lErr = 7: sErr = "Out of memory"
        Case ERROR_FILE_NOT_FOUND
            lErr = 53: sErr = "File not found"
        Case ERROR_PATH_NOT_FOUND
            lErr = 76: sErr = "Path not found"
        Case ERROR_BAD_FORMAT
            sErr = "The executable file is invalid or corrupt"
        Case SE_ERR_ACCESSDENIED
            lErr = 75: sErr = "Path/file access error"
        Case SE_ERR_ASSOCINCOMPLETE
            sErr = "This file type does not have a valid file association."
        Case SE_ERR_DDEBUSY
            lErr = 285: sErr = "The file could not be opened because the target application is busy.  Please try again in a moment."
        Case SE_ERR_DDEFAIL
            lErr = 285: sErr = "The file could not be opened because the DDE transaction failed.  Please try again in a moment."
        Case SE_ERR_DDETIMEOUT
            lErr = 286: sErr = "The file could not be opened due to time out.  Please try again in a moment."
        Case SE_ERR_DLLNOTFOUND
            lErr = 48: sErr = "The specified dynamic-link library was not found."
        Case SE_ERR_FNF
            lErr = 53: sErr = "File not found"
        Case SE_ERR_NOASSOC
            sErr = "No application is associated with this file type."
        Case SE_ERR_OOM
            lErr = 7: sErr = "Out of memory"
        Case SE_ERR_PNF
            lErr = 76: sErr = "Path not found"
        Case SE_ERR_SHARE
            lErr = 75: sErr = "A sharing violation occurred."
        Case Else
            sErr = "An error occurred occurred whilst trying to open or print the selected file."
        End Select
        Err.Raise lErr, , App.EXEName & ".GShell", sErr
        ShellEx = False
    End If
End Function

Sub DefinePosicaoECarregaForm(frmForm As Form, _
                     Optional bytModo As Byte, _
                     Optional strPosicao As String)
    Dim objObjeto   As Object
    PosicionaForm MDIMenu, frmForm, strPosicao
    For Each objObjeto In frmForm.Controls
        If LCase(Mid(objObjeto.Name, 1, 4)) = "cmd_" Then
            Call gblnCarregaIcone(objObjeto, "Mao", True)
        End If
    Next
    frmForm.Show bytModo
End Sub

'================================================================================'
'================================================================================'
'================================================================================'
'================================================================================'
'Funções criadas / alteradas por Reinato
'================================================================================'
'================================================================================'
'================================================================================'
'================================================================================'
Public Sub LeDaTabelaParaObj(strTabela As String, _
                             objObjeto As Object, _
                  ParamArray vntParametro())
    
    '--------------------------------------------------------------------------
    ' FUNÇÃO USADA PARA LER DA TABELA 'BASE DE DADOS SQL' E PREENCHER OS
    ' OBJETOS NO FORMULÁRIO
    '--------------------------------------------------------------------------
    ' PARÂMETROS:
    ' 1 - strTabela(tabela de onde será lido os dados
    ' 2 - objObjeto(lista a ser preenchida 'ListView, CombBox, ETC' ou
    '               formulário 'Form' onde serão mostrados as informações
    ' 3 - vntParametro(pode conter várias informações opcionais tais como:
    '     a) um valor Boolean indicando se mostra a barra de progressão
    '        com o botão para cancelar
    '     b) uma query específica. ex: um select com várias tabelas
    '     c) uma string com uma condição específica, ou seja, diferente da PKId
    '     d) etc.
    '--------------------------------------------------------------------------
    
'******************************************************************************************
' Data: 10/06/2003
' Alteração: - Alterada instrução IF que determinava se o comando SQL a ser executado é uma
'            stored procedure.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim blnNaoHabilita  As Boolean
    Dim strCondicao     As String
    Dim strQuery        As String
    Dim lngInd          As Long
    Dim vntItem         As Variant
    Dim strSql          As String
    Dim objList         As Object
    Dim objControl      As Object
    Dim vntAux          As Variant
    Dim strCampo        As String
    Dim objCampo        As ADODB.Field
    Dim adoResultado    As ADODB.Recordset
    
    Dim objFilho As Object
    Dim objContainer As Object
    
    Screen.MousePointer = vbHourglass
    
    If TypeOf objObjeto Is ListView Then
        objObjeto.ListItems.Clear
    ElseIf TypeOf objObjeto Is ComboBox Then
        objObjeto.Clear
    ElseIf TypeOf objObjeto Is DataCombo Or TypeOf objObjeto Is DataList Then
        objObjeto.RowSource = Nothing
        objObjeto.BoundText = ""
    ElseIf TypeOf objObjeto Is Form Then
        'Verifica se foi informado a query específica ou condição especial
        For Each vntItem In vntParametro
            If InStr(UCase(vntItem), "SELECT") <> 0 Then
                strQuery = Trim(vntItem) 'query específica
            ElseIf UCase(vntItem) = "TRUE" Then
                blnNaoHabilita = vntItem
            ElseIf UCase(vntItem) <> "FALSE" Then
                lngInd = lngInd + 1
        'Condição especial da seginte forma:
        '1º coluna 'EX. strCodigo
        '2º valor 'EX. 20 (vinte)
        'Sempre nessa ordem (coluna1, valor1, coluna2, valor2....
                If lngInd Mod 2 <> 0 Then 'verifica se é coluna
                    If Trim(strCondicao) = "" Then
                        strCondicao = "WHERE " & Trim(vntItem) & " = "
                    Else
                        strCondicao = strCondicao & "AND " & Trim(vntItem) & " = "
                    End If
                Else
                    strCondicao = strCondicao & "'" & Trim(vntItem) & "' "
                End If
            End If
        Next
        strSql = ""
        If Trim(strQuery) <> "" Then
            strSql = strQuery
        ElseIf Trim(strCondicao) <> "" Then
            strSql = strSql & "SELECT * FROM " & Trim(strTabela) & " "
            strSql = strSql & strCondicao
        Else
            strSql = strSql & "SELECT * FROM " & Trim(strTabela) & " "
            strSql = strSql & gstrCondicaoGeral(objObjeto)
        End If
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            If adoResultado.EOF = False Then
                'Loop eliminado pois o nome do objeto do form é igual ao do campo da tabela
                'sem o prefixo - EX: txtstrDescricao = strDescricao
                For Each objControl In objObjeto.Controls
                    'Condição substituída
                    'Se o nome do objeto do form contém "_" na quarta posição é um objeto auxiliar e não
                    'entra na atualização, assim como label´s.
                    If Mid(objControl.Name, 4, 1) <> "_" And (Not TypeOf objControl Is Label) Then
                        'Procura no nome do objeto do form o nome do campo na tabela
                        strCampo = Right(objControl.Name, Len(objControl.Name) - 3)
                        'Faz a atribuição
                        
                        AtribuiValorDoSql objControl, adoResultado.Fields(strCampo)
                    ElseIf TypeOf objControl Is TDBGrid Then
                        gCorLinhaSelecionada objControl
                    End If
                Next
            Else
                For Each objControl In objObjeto.Controls
                    If TypeOf objControl Is TDBGrid Then
                        If objControl.FilterActive = True Then
                            LimpaObjeto objObjeto
                            Screen.MousePointer = vbNormal
                            Exit Sub
                        End If
                    End If
                Next
                ExibeMensagem "Este registro foi excluído por outro usuário."
                Screen.MousePointer = vbNormal
                Exit Sub
            End If
        End If
        If blnNaoHabilita = False Then
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
        End If
        Screen.MousePointer = vbNormal
        Exit Sub
    ElseIf TypeOf objObjeto Is Frame Then
        For Each vntItem In vntParametro
            If InStr(UCase(vntItem), "SELECT") <> 0 Then
                strQuery = Trim(vntItem)
            ElseIf UCase(vntItem) = "TRUE" Then
                blnNaoHabilita = vntItem
            ElseIf UCase(vntItem) <> "FALSE" Then
                lngInd = lngInd + 1
                If lngInd Mod 2 <> 0 Then
                    If Trim(strCondicao) = "" Then
                        strCondicao = "WHERE " & Trim(vntItem) & " = "
                    Else
                        strCondicao = strCondicao & "AND " & Trim(vntItem) & " = "
                    End If
                Else
                    strCondicao = strCondicao & "'" & Trim(vntItem) & "' "
                End If
            End If
        Next
        strSql = ""
        If Trim(strQuery) <> "" Then
            strSql = strQuery
        ElseIf Trim(strCondicao) <> "" Then
            strSql = strSql & "SELECT * FROM " & Trim(strTabela) & " "
            strSql = strSql & strCondicao
        Else
            strSql = strSql & "SELECT * FROM " & Trim(strTabela) & " "
            strSql = strSql & gstrCondicaoFrame(objObjeto)
        End If
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            If adoResultado.EOF = False Then
                For Each objFilho In objObjeto.Parent.Controls
                
                    Set objContainer = objFilho.Container
                    
                    If (objContainer.Name) = objObjeto.Name And objFilho.Name <> objObjeto.Name And Not TypeOf objFilho Is Label Then
                    
                        If Not TypeOf objFilho Is Label And objFilho.Tag <> "*" Then
                            strCampo = Right(objFilho.Name, Len(objFilho.Name) - 4)
                            AtribuiValorDoSql objFilho, adoResultado.Fields(strCampo)
                        ElseIf TypeOf objFilho Is TDBGrid Then
                            gCorLinhaSelecionada objFilho
                        End If
                    End If
                    
                    Set objContainer = Nothing
                    
                Next
            Else
                For Each objFilho In objObjeto.Parent.Controls
                    If TypeOf objFilho Is TDBGrid Then
                        If objFilho.FilterActive = True Then
                            LimpaObjeto objObjeto
                            Screen.MousePointer = vbNormal
                            Exit Sub
                        End If
                    End If
                Next
                ExibeMensagem "Este registro foi excluído por outro usuário."
                Screen.MousePointer = vbNormal
                Exit Sub
            End If
        End If
        If blnNaoHabilita = False Then
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
        End If
        Screen.MousePointer = vbNormal
        Exit Sub

    ElseIf TypeOf objObjeto Is TrueOleDBGrid70.Column Then
        'objobjeto.value =
        
        objObjeto.Value = vntParametro(1).Columns("strDescricao").Value
        Screen.MousePointer = vbNormal
        Exit Sub
    ElseIf TypeOf objObjeto Is TDBGrid = False And TypeOf objObjeto Is TDBDropDown = False Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    For Each vntItem In vntParametro
        If InStr(UCase(vntItem), "SELECT") <> 0 Then
            strQuery = Trim(vntItem)
'        ElseIf UCase(Mid(vntItem, 1, 3)) = "SP_" Then
        ElseIf ((UCase(Mid(vntItem, 1, 3)) = "SP_") And (bytDBType = EDatabases.SQLServer)) Or _
            ((InStr(1, vntItem, "{ call", vbTextCompare) > 0) And (bytDBType = EDatabases.Oracle)) Then
            strQuery = Trim(vntItem)
        ElseIf InStr(UCase(vntItem), "WHERE") <> 0 Then
            strCondicao = Trim(vntItem)
        ElseIf Trim(strSql) = "" Then
            If Trim(vntItem) <> "" Then
                strSql = "SELECT " & vntItem
            End If
        ElseIf Trim(vntItem) <> "" Then
            strSql = strSql & ", " & vntItem
        End If
    Next
    If Trim(strQuery) = "" And Trim(strSql) = "" Then
        strSql = gstrQuery(strTabela)
    ElseIf Trim(strQuery) <> "" Then
        strSql = strQuery
    Else
        strSql = strSql & " FROM " & strTabela
    End If
    If Trim(strCondicao) <> "" Then
        strSql = strSql & Chr(vbKeySpace) & strCondicao
    End If
    If Trim(strSql) = "" Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 80, adoResultado) Then
        If TypeOf objObjeto Is TDBGrid Or TypeOf objObjeto Is TDBDropDown Then
            Set objObjeto.DataSource = adoResultado
            objObjeto.Refresh
        ElseIf TypeOf objObjeto Is DataCombo Or TypeOf objObjeto Is DataList Then
            objObjeto.ListField = adoResultado.Fields(1).Name
            objObjeto.BoundColumn = adoResultado.Fields(0).Name
            Set objObjeto.RowSource = adoResultado
            
            If Not adoResultado.EOF Then 'Nino
                If objObjeto.HelpContextID = 1 Then
                    objObjeto.BoundText = adoResultado(0)
                End If
            End If
        Else
            Do While adoResultado.EOF = False
                If TypeOf objObjeto Is ComboBox Then
                    objObjeto.AddItem gstrENulo(adoResultado.Fields(1))
                    objObjeto.ItemData(objObjeto.NewIndex) = adoResultado.Fields(0)
                ElseIf TypeOf objObjeto Is ListView Then
                    lngInd = 0
                    For Each objCampo In adoResultado.Fields
                        If lngInd = 1 Then
                            Set objList = objObjeto.ListItems.Add(, , gvntValorParaObjLista(objCampo))
                        ElseIf lngInd <> 0 Then
                            objList.SubItems(lngInd - 1) = gvntValorParaObjLista(objCampo)
                        End If
                        lngInd = lngInd + 1
                    Next
                    objList.Tag = adoResultado(0)
                End If
                adoResultado.MoveNext
            Loop
        End If
        Set gobjBanco = Nothing
    End If

ErroLeTabelaEPreencheObj:
    Screen.MousePointer = vbNormal
    If Err <> 0 Then
        ExibeDetalheErro ""
        Resume FimLeTabelaEPreencheObj
    End If
FimLeTabelaEPreencheObj:

End Sub

Public Function gvntUltimoDiaDoMes(strData As String) As Variant
    Select Case Month(strData)
    Case 1, 3, 5, 7, 8, 10, 12
        gvntUltimoDiaDoMes = 31
    Case 4, 6, 9, 11
        gvntUltimoDiaDoMes = 30
    Case 2
        If Year(strData) Mod 4 = 0 Then
            gvntUltimoDiaDoMes = 29
        Else
            gvntUltimoDiaDoMes = 28
        End If
    End Select
End Function

Public Function gvntVrFormatoEspecifico(objObjeto As Object) As Variant
    
'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strMascara          As String
    Dim STRNOME             As String
    Dim vntValor            As Variant
    Dim intInd              As Integer
    Dim intLenMascara       As Integer
    Dim strSql              As String
    Dim adoResultado        As ADODB.Recordset
    STRNOME = Trim(UCase(objObjeto.Name))
    gvntVrFormatoEspecifico = objObjeto
    vntValor = gstrValorSemMascara(objObjeto)
    strMascara = ""
    strSql = ""
    strSql = strSql & "SELECT strMascara "
    strSql = strSql & "FROM "
    strSql = strSql & gstrConfiguracao & " "
    strSql = strSql & "WHERE "
'    strSql = strSql & "UPPER(SUBSTRING(strColuna, 4, LEN(strColuna))) = '"
    strSql = strSql & "UPPER(" & strSUBSTRING & "(strColuna, 4, " & strLen & "(strColuna))) = '"
    strSql = strSql & Mid(STRNOME, 7) & "'"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            strMascara = Trim(adoResultado!strMascara)
            For intInd = 1 To Len(Trim(strMascara))
                If Mid(strMascara, intInd, 1) = "0" Then
                    intLenMascara = intLenMascara + 1
                End If
            Next
            If Len(gstrENulo(vntValor)) < 15 Then
                intLenMascara = 15 - Len(gstrENulo(vntValor))
            ElseIf Len(gstrENulo(vntValor)) > 15 Then
                intLenMascara = Len(gstrENulo(vntValor)) - 15
            Else
                intLenMascara = 0
            End If
            gvntVrFormatoEspecifico = Trim(vntValor) & String(intLenMascara, "0")
        End If
    End If
End Function

Public Function gvntConvFormatoEspecificoParaSQL(vntObjeto As Variant, _
                                        Optional bytObjeto As Byte) As Variant
    Dim strMascara          As String
    Dim STRNOME             As String
    Dim vntValor            As Variant
    Dim intInd              As Integer
    Dim intLenMascara       As Integer
    If Trim(gstrMascaraContaContabil) = "" Then
        LeMascacaraEspecifica
    End If
    Select Case UCase(TypeName(vntObjeto))
    Case "TEXTBOX", "COMBOBOX", "FIELD"
        STRNOME = vntObjeto.Name
    End Select
    If InStr(UCase(STRNOME), "CONTACONTABIL") Or bytObjeto = 1 Then
        strMascara = gstrMascaraContaContabil
    ElseIf InStr(UCase(STRNOME), "CODIGOORCAMENTARIO") Or bytObjeto = 2 Then
        strMascara = gstrMascaraCodigoOrcamentario
    ElseIf InStr(UCase(STRNOME), "ELEMENTODESPESA") Or bytObjeto = 3 Then
        strMascara = gstrMascaraElementoDespesa
    End If
    vntValor = gstrValorSemMascara(vntObjeto)
    gvntConvFormatoEspecificoParaSQL = vntObjeto
    If Trim(strMascara) <> "" Then
        For intInd = 1 To Len(Trim(strMascara))
            If Mid(strMascara, intInd, 1) = "0" Then
                intLenMascara = intLenMascara + 1
            End If
        Next
        If Len(gstrENulo(vntValor)) < 15 Then
            intLenMascara = 15 - Len(gstrENulo(vntValor))
        ElseIf Len(gstrENulo(vntValor)) > 15 Then
            intLenMascara = Len(gstrENulo(vntValor)) - 15
        Else
            intLenMascara = 0
        End If
        gvntConvFormatoEspecificoParaSQL = Trim(vntValor) & String(intLenMascara, "0")
    End If
End Function

Public Sub VirificaGradeListView(frmForm As Form, _
                        Optional blnAlterando As Boolean)
    
    '--------------------------------------------------------------------
    ' SUB USADA PARA MUDAR ESTILO DE GRADE NOS OBJETOS TAIS COMO,
    ' Grid COLOCANDO OU RETIRANDO AS LINHAS DE GRADE
    '--------------------------------------------------------------------
    ' PARÂMETRO:
    '
    ' 1 - frmForm (Formulário que está sendo carregado (chamado)
    '--------------------------------------------------------------------
    
    Dim objControl  As Object
    For Each objControl In frmForm.Controls
        If TypeOf objControl Is TDBGrid Then
            If gblnListViewComGrade Then
                objControl.RowDividerStyle = 3
            Else
                objControl.RowDividerStyle = 0
            End If
        ElseIf TypeOf objControl Is ListView Then
            If gblnListViewComGrade Then
                objControl.GridLines = True
            Else
                objControl.GridLines = False
            End If
        ElseIf TypeOf objControl Is Label Then
            If objControl.Tag = 1 Then
                TrocaCorObjeto objControl, blnAlterando
            End If
        ElseIf TypeOf objControl Is TextBox _
        Or TypeOf objControl Is ComboBox _
        Or TypeOf objControl Is DataCombo Then
            If objControl.OLEDragMode = vbAutomatic Then
                TrocaCorObjeto objControl, blnAlterando
            End If
        End If
    Next
End Sub
    
Public Sub PadronizaToolBarRelatorio(rptRelatorio As Object, _
                                     Optional lblExercicio As Object, _
                                     Optional vntExercicio As Variant)
    
Dim objField As Object
    
    On Error GoTo ErroPadronizacao
    If gblnRestartRelatorio Then
        If intOrientacao <> 0 Then
            rptRelatorio.Printer.Orientation = intOrientacao
        End If
        Exit Sub
    End If
    Select Case rptRelatorio.Version
        Case "2.0"
            'Cláudio
'            With rptRelatorio.Sections("PageFooter").Controls
'                Set objfield = .Add("DDActiveReports2.Field")
'                objfield.Name = "txtNumPaginaTotal"
'                objfield.Visible = False
'                objfield.Style = "font-size: 8pt; text-align: center; "
'                objfield.SummaryRunning = 0
'                objfield.SummaryType = 4
'                .Item("lblPagina").Width = rptRelatorio.Width
'                .Item("lblPagina").Alignment = 2
'            End With
            
            With rptRelatorio.Toolbar
            
            
                .Tools.Item(0).Visible = False
                .Tools.Item(1).Visible = False
                .Tools.Item(3).Visible = False
                .Tools.Item(4).Visible = False
                .Tools.Item(5).Visible = False
                
                '.Tools.Item(6).Visible = False
                .Tools.Item(8).Visible = False
                .Tools.Item(9).Visible = False
                .Tools.Item(10).Visible = False
                
                .Tools.Item(2).Caption = ""
                .Tools.Item(2).Tooltip = "Imprimir"
                .Tools.Item(6).Tooltip = "Localizar"
                .Tools.Item(11).Tooltip = "Diminuir"
                .Tools.Item(12).Tooltip = "Aumentar"
                .Tools.Item(15).Tooltip = "Página anterior"
                .Tools.Item(16).Tooltip = "Página seguinte"
                .Tools.Item(17).Tooltip = "Número da Página"
                .Tools.Item(19).Caption = ""
                .Tools.Item(19).Tooltip = "Histórico anterior"
                .Tools.Item(20).Caption = ""
                .Tools.Item(20).Tooltip = "Histórico seguinte"
                .Tools.Add "&Exportar..."
                .Tools.Item(21).ID = 15
                .Tools.Item(21).Tooltip = "Exportar"
                .Tools.Add "&Fechar"
                .Tools.Item(22).ID = 14
                .Tools.Item(22).Tooltip = "Fechar"
                .Tools.Insert 3, "&Configurar..."
                .Tools.Item(3).ID = 16
                .Tools.Item(3).Tooltip = "Configurar impressão"
            End With
        Case Else
            With rptRelatorio.Toolbar
                .Tools.Item(0).Visible = False
                .Tools.Item(1).Visible = False
                .Tools.Item(2).Caption = ""
                .Tools.Item(2).Tooltip = "Imprimir"
                .Tools.Item(4).Tooltip = "Diminuir"
                .Tools.Item(5).Tooltip = "Aumentar"
                .Tools.Item(8).Tooltip = "Página anterior"
                .Tools.Item(9).Tooltip = "Página seguinte"
                .Tools.Item(10).Tooltip = "Número da Página"
                .Tools.Item(12).Caption = ""
                .Tools.Item(12).Tooltip = "Histórico anterior"
                .Tools.Item(13).Caption = ""
                .Tools.Item(13).Tooltip = "Histórico seguinte"
                .Tools.Add "&Exportar..."
                .Tools.Item(14).ID = 15
                .Tools.Item(14).Tooltip = "Exportar"
                .Tools.Add "&Fechar"
                .Tools.Item(15).ID = 14
                .Tools.Item(15).Tooltip = "Fechar"
                .Tools.Insert 3, "&Configurar..."
                .Tools.Item(3).ID = 16
                .Tools.Item(3).Tooltip = "Configurar impressão"
            End With
    End Select
    With rptRelatorio
        .Zoom = 77
    End With
    If IsMissing(vntExercicio) = False Then
        lblExercicio = "Exercício: " & Val(vntExercicio)
    ElseIf gblnProposta Then
        If UCase(App.ProductName) = "ORCAMENTARIO" Then
            lblExercicio = "Exercício: " & gintExercicio + 1
        Else
            lblExercicio = "Exercício: " & Year(gstrDataDoSistema) + 1
        End If
        'Nino
        'gblnProposta = False
    Else
        If UCase(App.ProductName) = "ORCAMENTARIO" Then
            lblExercicio = "Exercício: " & gintExercicio
        Else
            lblExercicio = "Exercício: " & Year(gstrDataDoSistema)
        End If
    End If
    Exit Sub
    
ErroPadronizacao:
    Resume FimPadronizacao
    
FimPadronizacao:
End Sub

Sub ProcuraRegistroGravado(objLista As Object, _
                           frmForm As Form, _
                           strTabela As String, _
                           blnAlterando As Boolean)
    Dim adoResultado    As ADODB.Recordset
    Dim strPKId         As String
    If blnAlterando Then
        strPKId = Val(frmForm.txtPKId)
    Else
        strPKId = glngPegaUltimaChave(strTabela, "PKID")
    End If
    If Trim(strPKId) <> "" Then
        If TypeOf objLista Is TDBGrid Then
            Set adoResultado = objLista.DataSource
            adoResultado.Find "PKId = '" & strPKId & "'"
            objLista.MarqueeStyle = dbgHighlightRow
        ElseIf TypeOf objLista Is ListView Then
            Call gblnEncontroItemNoListView(objLista, strPKId)
        End If
    End If
End Sub

Public Function SalvarGeral(strTabela As String, strModoOperacao As String, frmForm As Form, objLista As Object, _
                            strQuery As String, Optional LimpaForm As Boolean = True, Optional blnAlterando As Boolean, Optional blnMsg As Boolean = False) As Boolean
'======================================================================================================
'Função monta query para manutenções no banco chamando a dll para tratamento de erros ou execução
'da query
'Parâmetros:
    'strTabela       => A tabela que recebera manutenção
    'strModoOperacao => (I)nclusão, (A)lteração, (E)xclusão
    'frmForm         => O formulário que trata a manutenção na tabela
'======================================================================================================

'******************************************************************************************
' Data: 07/03/2003
' Alteração: - Quando o Banco de Dados corrente for o Oracle e o tipo de dados a ser
'              inserido é um DATE, este valor deve ser passado para a função
'              gstrFormataDataOracle.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim blnErro As Boolean
Dim adoResultado            As ADODB.Recordset
Dim strSql                  As String
Dim strLabel                As String
Dim strFinalString          As String
'Auxiliares para o loop no formulário
Dim i As Integer
Dim j As Integer

'Possui 4 colunas
    '0 => O nome do campo da tabela (strTabela)
    '1 => O valor do campo do formulário ("NULL" quando vazio)
    '2 => O título (Label) do campo no formulário
    '3 => O prefixo do objeto ("txt","chk","cbo",etc...)
    
'Tag de objetos com nº 1 => TextBox que utilizam máscara específica e guardam na tabela o valor
                            'Sem a máscara
Dim vetCampos() As String

'Auxiliar para chamar a DLL e tratar o retorno da mesma
Dim strAux As String

'A DLL que fará o tratamento das manutenções no banco de dados
Dim objFuncao As clsFuncoes

'Atribui o campo descrição do objeto que identifica-se no form
Dim strDescricao As String

On Error GoTo Err_MontaValores
    
    With frmForm
        For i = 0 To .Controls.Count - 1
            If UCase(.Controls(i).Name) = "TXTSTRDESCRICAO" Or UCase(.Controls(i).Name) = "TXTSTRNOME" Then
                strDescricao = .Controls(i).Text
                Exit For
            End If
        Next
    End With

    If blnMsg = True Then
        GoTo IguinoraConfirmacao
    End If
    
    If gblnExclusaoGravacaoOk(strModoOperacao, strDescricao) Then
    
IguinoraConfirmacao:

        On Error Resume Next
        If Not frmForm.blnDadosOk Then
            If Err.Number = 0 Then SalvarGeral = False: Exit Function
        End If
        On Error GoTo Err_MontaValores
        
        'Cria uma nova instância da DLL
        Set objFuncao = New clsFuncoes
        
        'Primeiro índice do vetor (linha)
        j = 0
        
        With frmForm
            'Exclusão
            If strModoOperacao = "E" Then
                ReDim vetCampos(1, 0)
                vetCampos(0, 0) = "PKID"
                vetCampos(1, 0) = Val(.Controls("txtpkid").Text)
                strDescricao = .Controls("txtstrDescricao").Text
            Else
                'Percorre os controles do form
                For i = 0 To .Controls.Count - 1
                    
                    'Elimina os objetos que não tem relacionamento com os campos da tabela
                    'Estes controles são identificados por ter seu prefixo (tres letras) separados do
                    'nome por "_"
                    If Mid(.Controls(i).Name, 4, 1) <> "_" And (Not TypeOf .Controls(i) Is Label) Then
                        
                        If Not (TypeOf .Controls(i) Is OptionButton) Or .Controls(i) = True Then
                        
                            'If UCase(.Controls(i).Name) = "TXTSTRDESCRICAO" Or UCase(.Controls(i).Name) = "TXTSTRNOME" Then
                            '    strDescricao = .Controls(i).Text
                            'End If
                            
                            'Redimensiona o Vetor para o novo campo encontrado
                            ReDim Preserve vetCampos(3, j)
                            
                            'Elimina o prefixo do objeto do formulário para se descobrir o nome do campo na tabela
                            'ex:
                                'Nome do Campo no Form  = Nome do Campo na Tabela
                                'txtStrDescricao        = StrDescricao
                            strAux = Right(.Controls(i).Name, Len(.Controls(i).Name) - 3)
                            
                            'Atribui ao vetor o nome do campo na tabela
                            vetCampos(0, j) = strAux
                            
                            'Verifica se controle está preenchido
                            If Trim(.Controls(i)) = "" Then
                                vetCampos(1, j) = "NULL"
                            Else
                                If TypeOf .Controls(i) Is DTPicker Then
        '                            vetCampos(1, j) = Format(.Controls(i), "YYYY/MM/DD")
                                    If bytDBType = EDatabases.SQLServer Then
                                        vetCampos(1, j) = Format(.Controls(i), "YYYY/MM/DD")
                                    ElseIf bytDBType = EDatabases.Oracle Then
                                        vetCampos(1, j) = gstrFormataDataOracle(.Controls(i), "yyyy/mm/dd")
                                    End If
                                ElseIf TypeOf .Controls(i) Is MaskEdBox Then
                                    If UCase(Left(strAux, 3)) = "DTM" Then
                                        If Trim(.Controls(i).FormattedText) <> "/  /" Then
                                            If gblnDataValida(.Controls(i).FormattedText, True) Then
                                                'if len(.Controls(i).FormattedText)
        '                                        vetCampos(1, j) = Format(.Controls(i).FormattedText, "YYYY/MM/DD hh:mm:ss")
                                                If bytDBType = EDatabases.SQLServer Then
                                                    vetCampos(1, j) = Format(.Controls(i).FormattedText, "YYYY/MM/DD hh:mm:ss")
                                                ElseIf bytDBType = EDatabases.Oracle Then
                                                    vetCampos(1, j) = gstrFormataDataOracle(.Controls(i).FormattedText)
                                                End If
                                            Else
                                                Exit Function
                                            End If
                                        Else
                                            vetCampos(1, j) = "NULL"
                                        End If
                                    ElseIf InStr(1, strAux, "CNPJCPF") > 0 Then
                                        If Len(Trim(gstrValorSemMascara(gstrCGCCPFFormatado(.Controls(i).ClipText)))) = 11 Then
                                            If gblnCPFOk(.Controls(i).ClipText) Then
                                                vetCampos(1, j) = gstrValorSemMascara(.Controls(i).ClipText)
                                            Else
                                                ExibeMensagem "CPF inválido"
                                                If .Controls(i).Enabled Then
                                                    .Controls(i).SetFocus
                                                End If
                                                Exit Function
                                            End If
                                        ElseIf Len(Trim(gstrValorSemMascara(.Controls(i).ClipText))) = 14 Then
                                            If gblnCGCOk(.Controls(i).ClipText) Then
                                                vetCampos(1, j) = gstrValorSemMascara(.Controls(i).ClipText)
                                            Else
                                                ExibeMensagem "CNPJ inválido"
                                                If .Controls(i).Enabled Then
                                                    .Controls(i).SetFocus
                                                End If
                                                Exit Function
                                            End If
                                        Else
                                            ExibeMensagem "CNPJ/CPF inválido"
                                            If .Controls(i).Enabled Then
                                                .Controls(i).SetFocus
                                            End If
                                            Exit Function
                                        End If
                                    Else
                                        vetCampos(1, j) = .Controls(i).ClipText
                                    End If
                                ElseIf TypeOf .Controls(i) Is OptionButton Then
                                    vetCampos(1, j) = .Controls(i).Index
                                ElseIf TypeOf .Controls(i) Is DataCombo Then
                                    'Valor selecionado/digitado no DataCombo encontra-se na lista
                                    If .Controls(i).MatchedWithList Then
                                        'Guarda o código (BoundText) e o valor preenchido (Text) do DataCombo
                                        'que será mostrado ao usuário caso haja um erro de integridade referencial
                                        'com a tabela (duplicidade, etc...)
                                        
                                        'Campo na tabela do tipo texto, identificado por "str"
                                        If UCase(Mid(.Controls(i).Name, 4, 3)) = "STR" Then
                                            vetCampos(1, j) = .Controls(i).Text
                                        Else
                                            vetCampos(1, j) = .Controls(i).BoundText & "_" & .Controls(i).Text
                                        End If
                                    Else
                                        'Campo na tabela do tipo texto, identificado por "str"
                                        If UCase(Mid(.Controls(i).Name, 4, 3)) = "STR" Then
                                            'Guarda o valor preenchido (Text) do DataCombo
                                            vetCampos(1, j) = Trim(.Controls(i).Text)
                                        'Campo na tabela do tipo numérico, identificado por "int"
                                        ElseIf UCase(Mid(.Controls(i).Name, 4, 3)) = "INT" Then
                                            'Atribui-se um caracter junto a descrição do campo para que possa, no COM/DCOM,
                                            'padronizar mensagens de erro de conversão de tipos de dados.
                                            '(O usuário poderá digitar um número, aí teríamos outro tipo de erro como
                                            'Constraint, integridade de relacionamentos, duplicidade, etc...)
                                            'SELECT PKId, strDescricao FROM tblBairro ORDER BY strDescricao;strDescricao
                                            If .Controls(i).Tag <> "" Then
                                                strSql = Left(.Controls(i).Tag, InStr(.Controls(i).Tag, ",") - 1)
                                                
                                                If Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), " ")) = "" Then
                                                    strSql = strSql & " FROM " & Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), Len(.Controls(i).Tag))
                                                Else
                                                    strFinalString = Right(Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), " ") + 5), 6)
                                                    Select Case Trim(UCase(strFinalString))
                                                        Case "ORDER", "GROUP", "WHERE"
                                                            strSql = strSql & " FROM " & Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), " "))
                                                        Case Else
                                                            If Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), ",")) <> "" Then
                                                                strSql = strSql & " FROM " & Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl")), ","))
                                                            ElseIf Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), ";")) <> "" Then
                                                                strSql = strSql & " FROM " & Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), ";"))
                                                            Else
                                                                strSql = strSql & " FROM " & Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), Len(.Controls(i).Tag))
                                                            End If
                                                    End Select
                                                End If
                                                strSql = strSql & " WHERE " & Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStrRev(.Controls(i).Tag, ";"))
                                                strSql = strSql & " ='" & .Controls(i).Text & "'"
                        
                                                Set gobjBanco = New clsBanco
                                                
                                                If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                                                    If adoResultado.EOF Then
                                                        
                                                        On Error Resume Next
                                                        strLabel = .Controls(Replace(.Controls(i).Name, "dbc", "lbl")).Caption & " "
                                                        On Error GoTo 0
                                                        
                                                        ExibeMensagem "O valor digitado para o campo " & strLabel & "não está correto!"
            
                                                        .Controls(i).SetFocus
                                                        SalvarGeral = False
                                                        Exit Function
                                                    Else
                                                        vetCampos(1, j) = adoResultado(0) '"[" & Trim(.Controls(i).Text) & "]"
                                                    End If
                                                End If
                                                
                                                Set gobjBanco = Nothing
                                            Else
                                                On Error Resume Next
                                                strLabel = .Controls(Replace(.Controls(i).Name, "dbc", "lbl")).Caption & " "
                                                On Error GoTo 0
                                                            
                                                ExibeMensagem "O valor digitado para o campo " & strLabel & "não está correto!"
                
                                                If .Controls(i).Enabled Then
                                                    .Controls(i).SetFocus
                                                End If
                                                SalvarGeral = False
                                                Exit Function
                                            'vetCampos(1, j) = "[" & Trim(.Controls(i).Text) & "]"
                                            End If
                                        Else
                                            MsgBox "O campo " & Right(.Controls(i).Name, Len(.Controls(i).Name) - 3) & _
                                                   " na tabela " & strTabela & " contém um prefixo não identificado " & _
                                                   " de acordo com a padronização adotada no desenvovimento " & _
                                                   " de objetos COM/DCOM." & Chr(10) & Chr(10) & _
                                                   "Por favor, entre em contato com o CallCenter - 0800-00800", vbExclamation
                                            Exit Function
                                        End If
                                    End If
                                ElseIf TypeOf .Controls(i) Is ComboBox Then
                                    If .Controls(i).ListIndex >= 0 Then
                                        'Campo na tabela do tipo texto, identificado por "str"
                                        If UCase(Mid(.Controls(i).Name, 4, 3)) = "STR" Then
                                            vetCampos(1, j) = .Controls(i).Text
                                        Else
                                            vetCampos(1, j) = .Controls(i).ItemData(.Controls(i).ListIndex) & "_" & .Controls(i).Text
                                        End If
                                    Else
                                        'Campo na tabela do tipo texto, identificado por "str"
                                        If UCase(Mid(.Controls(i).Name, 4, 3)) = "STR" Then
                                            'Guarda o valor preenchido (Text) do ComboBox
                                            vetCampos(1, j) = .Controls(i).Text
                                        ElseIf UCase(Mid(.Controls(i).Name, 4, 3)) = "INT" Then
                                            'Atribui-se um caracter junto a descrição do campo para que possa, no COM/DCOM,
                                            'padronizar mensagens de erro de conversão de tipos de dados.
                                            '(O usuário poderá digitar um número, aí teríamos outro tipo de erro como
                                            'Constraint, integridade de relacionamentos, duplicidade, etc...)
                                            vetCampos(1, j) = "[" & Trim(.Controls(i).Text) & "]"
                                        Else
                                            MsgBox "O campo " & Right(.Controls(i).Name, Len(.Controls(i).Name) - 3) & _
                                                   " na tabela " & strTabela & " contém um prefixo não identificado " & _
                                                   " de acordo com a padronização adotada no desenvovimento " & _
                                                   " de objetos COM/DCOM." & Chr(10) & Chr(10) & _
                                                   "Por favor, entre em contato com o CallCenter - 0800-00800", vbExclamation
                                            Exit Function
                                        End If
                                    End If
                                Else
                                    If UCase(Left(strAux, 3)) = "DTM" Then
                                        If gblnDataValida(.Controls(i), True) Then
        '                                    vetCampos(1, j) = Format(.Controls(i), "YYYY/MM/DD HH:MM:SS")
                                            If bytDBType = EDatabases.SQLServer Then
                                                vetCampos(1, j) = Format(.Controls(i), "YYYY/MM/DD HH:MM:SS")
                                            ElseIf bytDBType = EDatabases.Oracle Then
                                                vetCampos(1, j) = gstrFormataDataOracle(.Controls(i))
                                            End If
                                        Else
                                            Exit Function
                                        End If
                                    ElseIf InStr(UCase(.Controls(i).Name), "INTCEP") > 0 Then
                                        vetCampos(1, j) = gstrValorSemMascara(.Controls(i))
                                    ElseIf InStr(UCase(.Controls(i).Name), "CNPJCPF") > 0 Or _
                                       InStr(UCase(.Controls(i).Name), "CNPJ") > 0 Then
                                        If Len(Trim(gstrValorSemMascara(.Controls(i)))) = 11 Then
                                            If gblnCPFOk(gstrValorSemMascara(.Controls(i))) Then
                                                vetCampos(1, j) = gstrValorSemMascara(.Controls(i))
                                            Else
                                                ExibeMensagem "CPF inválido"
                                                If .Controls(i).Enabled Then
                                                    .Controls(i).SetFocus
                                                End If
                                                Exit Function
                                            End If
                                        ElseIf Len(Trim(gstrValorSemMascara(.Controls(i)))) = 14 Then
                                            If gblnCGCOk(gstrValorSemMascara(.Controls(i))) Then
                                                vetCampos(1, j) = gstrValorSemMascara(.Controls(i))
                                            Else
                                                ExibeMensagem "CNPJ inválido"
                                                If .Controls(i).Enabled Then
                                                    .Controls(i).SetFocus
                                                End If
                                                Exit Function
                                            End If
                                        Else
                                            ExibeMensagem "CNPJ/CPF inválido"
                                            If .Controls(i).Enabled Then
                                                .Controls(i).SetFocus
                                            End If
                                            Exit Function
                                        End If
                                    ElseIf Val(.Controls(i).Tag) = 1 Then
                                    'Tag de objetos com nº 1 => TextBox que utilizam máscara
                                    'específica e guardam na tabela
                                    'o valor Sem a máscara
                                    '17/05/2001
                                        vetCampos(1, j) = gvntConvFormatoEspecificoParaSQL(.Controls(i))
                                    ElseIf IsNumeric(.Controls(i)) And _
                                           UCase(.Controls(i).Name) <> "TXTSTRMASCARA" And _
                                           Val(.Controls(i).Tag) <> 2 Then
                                        vetCampos(1, j) = gstrConvVrParaSql(.Controls(i))
                                    Else
                                        vetCampos(1, j) = .Controls(i)
                                    End If
                                End If
                            End If
                            
                            'Título do campo
                            '(Elimina o PKID - autonumeração - que será utilizado apenas nas cláusulas WHERE
                            'das query's de ação (não há título para este campo))
                            If UCase(.Controls(i).Name) <> "TXTPKID" Then
                                If TypeOf .Controls(i) Is OptionButton Then
                                    strAux = "opt" & strAux
                                    vetCampos(2, j) = .Controls(i).Caption
                                ElseIf TypeOf .Controls(i) Is CheckBox Then
                                    strAux = "chk" & strAux
                                    vetCampos(2, j) = .Controls(i).Caption
                                Else
                                    strAux = "lbl" & strAux
                                    vetCampos(2, j) = .Controls(strAux).Caption
                                    If blnErro Then
                                        strAux = Replace(strAux, "lbl", "fra_")
                                        vetCampos(2, j) = .Controls(strAux).Caption
                                    End If
                                End If
                            End If
                            'Prefixo do tipo do objeto
                            vetCampos(3, j) = Left(.Controls(i).Name, 3)
                            
                            'Incrementa o índice das linhas do vetor
                            j = j + 1
                        End If
                    End If
                    
                Next i
            End If
        End With
    'gcncADOMain.BeginTrans
    
        gblnCancelarInclusao = False
        'Chamada da DLL para tratamento e manutenção no banco
        strAux = objFuncao.bolFuncGravaDados(vetCampos, strModoOperacao, strTabela, gcncADOMain)
        'Diferente de "" a manutenção no banco de dados não foi bem sucedida
        If strAux <> "" Then
            'Mostra mensagem de erro encontrada e tratada também na DLL
            vetCampos = Split(strAux, "}")
            If UBound(vetCampos, 1) > 0 Then
                strAux = Trim(vetCampos(1))
                MsgBox "O campo " & frmForm.Controls(strAux).Caption & " não pode ser nulo!"
            Else
                MsgBox strAux
            End If
            'Foco no objeto com dados incorretos
            If objFuncao.strNomeCampo <> "" Then
                If Not (TypeOf frmForm.Controls(objFuncao.strNomeCampo) Is OptionButton) Then
                    If frmForm.Controls(objFuncao.strNomeCampo).Enabled Then
                        frmForm.Controls(objFuncao.strNomeCampo).SetFocus
                    End If
                End If
            End If
            SalvarGeral = False
            'gcncADOMain.RollbackTrans
        Else
            If objLista Is Nothing = False And gblnListagemAutomatica Then
                LeDaTabelaParaObj strTabela, objLista, strQuery
                ProcuraRegistroGravado objLista, frmForm, strTabela, blnAlterando
            End If
            If LimpaForm Then
                LimpaObjeto frmForm 'Não está limpando Natureza Jurídica e Data do cadastro
            End If
            SalvarGeral = True
            blnAlterando = False
        '    gcncADOMain.CommitTrans
        End If
    Else
    '    gcncADOMain.RollbackTrans
        gblnCancelarInclusao = True
    End If



Exit Function
'Erro imprevisto (DEFINIR MELHOR MENSAGEM)
Err_MontaValores:
If Err.Number = 730 Then
    blnErro = True
    Resume Next
ElseIf Err.Number <> 438 Then 'Objeto não suporta setfocus
    ExibeDetalheErro ""
End If
End Function

'ToolBarGeral strModoOperacao, gstrGrupoMaterialServico, mblnAlterando, tdb_GrpMatSer, Me, tdb_GrpMatSer, strSql, strSql
Public Function ToolBarGeral(strModoOperacao As String, _
                             strTabela As String, _
                             blnAlterando As Boolean, _
                    Optional objLista As Object, _
                    Optional frmForm As Form, _
                    Optional objGeral As Object, _
                    Optional strQueryRefresh As String, _
                    Optional strQueryAplicar As String, _
                    Optional objRelatorio As Object, _
                    Optional strQueryRelatorio As String, _
                    Optional blnLimpaForm As Boolean = True) As Boolean

    '---------------------------------------------------------------------------
    ' SUB USADA PARA CHAMAR AS FUNÇÕES BÁSICAS DA BARRA DE FERRAMENTAS
    '---------------------------------------------------------------------------
    ' PARÂMETROS:
    ' 1 - bttBotao(botão que foi clicado)
    ' 2 - strTabela(tabela onde será efetuada a operação)
    ' 3 - blnAlterando(flag indicando se está fazendo alteração ou incluindo)
    ' 4 - objLista(objeto onde será mostrado os dados 'ex.listview')
    ' 5 - frmForm(formulário de onde está sendo chamada a sub)
    ' 6 - objGeral(objeto que será utilizado para preencher lista de dados, caso,
    '              tenha que retornar informações para outro formulário)
    ' 7 - strQueryRefresh(Query que será executada caso a lista de dados seja
    '                     com base numa query específica)
    ' 8 - strQueryAplicar(query que será executada caso a lista que retornará para
    '                     outro formulário ser uma query específica)
    ' 9 - objRelatorio(active report)
    ' 10 - strQueryRelatorio(query que será usada para gerar o relatório)
    '---------------------------------------------------------------------------
    Dim strSql As String
    
    Select Case UCase(strModoOperacao)
        Case gstrNovo
            LimpaObjeto frmForm, blnAlterando
            ToolBarGeral = True
        Case gstrSalvar
            'If Not gblnDadosOk(frmForm) Then Exit Function
            ToolBarGeral = SalvarGeral(strTabela, IIf(blnAlterando, "A", "I"), frmForm, objLista, strQueryRefresh, blnLimpaForm, blnAlterando)
        Case gstrDeletar
            ToolBarGeral = SalvarGeral(strTabela, "E", frmForm, objLista, strQueryRefresh, blnLimpaForm, blnAlterando)
        Case gstrAplicar
            AplicarGeral frmForm, objGeral, objLista, strTabela, , strQueryAplicar
        Case gstrGrade
            MudaGradeDBGrid frmForm
        Case gstrImprimir
            ImprimeRelatorio objRelatorio, strQueryRelatorio
        Case gstrRefresh
            
            LeDaTabelaParaObj strTabela, objLista, strQueryRefresh
        Case gstrLocalizar
            If strQueryRefresh = "" Then
                strQueryRefresh = "SELECT * FROM " & strTabela
            End If
            strSql = LocalizarGeral(frmForm, strQueryRefresh, strTabela)
            If strSql <> "" Then
                LeDaTabelaParaObj strTabela, objLista, strSql
            End If
        Case gstrPreencherLista
            PreencherListaDeOpcoes frmForm.ActiveControl
        Case gstrFechar
            Unload frmForm
    End Select
End Function

Private Function LocalizarGeral(frmForm As Form, ByVal strSql As String, ByVal strTabela As String) As String

'******************************************************************************************
' Data: 13/03/2003
' Alteração: - Alterada a instrução IF na montagem dinâmica da cláusula de modo que se o
'            campo contiver 4000 caracteres este não é incluído na cláusula, devido ao
'            motivo de que o Oracle não permite um conjunto de mais do que 4000 caracteres
'            nas cláusulas WHERE, INSERT e UPDATE
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 14/03/2003
' Alteração: - Alterada a instrução que incluía campos data na cláusula WHERE.
'            - Modificada a procura do alias da tabela.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim i                   As Integer
Dim intPos              As Integer
Dim strCondicao         As String
Dim strNomeCampo        As String
Dim strOrdena           As String
Dim strGroup            As String
Dim strWhere            As String
Dim strApelido          As String
Dim strAux              As String
Dim blnAchou            As Boolean
Dim strTextTemp         As String
Dim strFinalString      As String
Dim strSQLNotMatched    As String
Dim adoResultado        As ADODB.Recordset
Dim strLabel            As String
'Dim vetApelido() As String

''Comentado por MARCO
''intPos = InStr(1, strSql, strTabela & " AS ")
''
''strApelido = strTabela
''
''If intPos > 0 Then
''
''Else
''    intPos = InStr(1, strSql, strTabela)
''    If intPos > 0 Then
''        If InStr(intPos, strSql, ",") > 0 Then
''
''            For i = intPos To Len(strSql)
''
''                If Mid(strSql, i, 1) = " " Then
''                    blnAchou = True
''                End If
''
''                If Mid(strSql, i, 1) = "," Then
''                    Exit For
''                End If
''
''                If blnAchou Then
''                    strAux = strAux & Mid(strSql, i, 1)
''                End If
''            Next i
''
''            strApelido = strAux & " "
''
''        End If
''    End If
''    If strTabela = "tblLogradouro" Then
''        strApelido = "L"
''    End If
''End If
'strApelido = strTabela
strApelido = gstrGetTableAlias(strSql, strTabela)
'vetApelido = Split(Trim(strSQL), " ")
'For i = 0 To UBound(vetApelido)
'    If UCase(vetApelido(i)) = UCase(Trim(strTabela)) Then
'        If i < UBound(vetApelido) Then
'            If UCase(Trim(vetApelido(i + 1))) = "AS" Then
'                strApelido = UCase(Trim(vetApelido(i + 2)))
'                intPos = InStr(1, strApelido, ",")
'                If intPos > 0 Then
'                    strApelido = Mid(strApelido, 1, intPos - 1)
'                End If
'                Exit For
'            End If
'        End If
'    End If
'Next

Screen.MousePointer = vbHourglass

intPos = InStrRev(UCase(strSql), "ORDER")

If intPos > 0 Then
    strOrdena = Right(strSql, (Len(strSql) - intPos) + 1)
    strSql = Left(strSql, intPos - 1)
End If

intPos = InStrRev(UCase(strSql), "GROUP")

If intPos > 0 Then
    strGroup = Right(strSql, (Len(strSql) - intPos) + 1)
    strSql = Left(strSql, intPos - 1)
End If

intPos = InStrRev(UCase(strSql), "WHERE")

If intPos > 0 Then
    strWhere = Right(strSql, (Len(strSql) - intPos) + 1)
    strSql = Left(strSql, intPos - 1)
End If

With frmForm
    For i = 0 To .Controls.Count - 1
        If Mid(.Controls(i).Name, 4, 1) <> "_" And (Not TypeOf .Controls(i) Is Label) Then
            strNomeCampo = Mid(.Controls(i).Name, 4, Len(.Controls(i).Name))
'            If TypeOf .Controls(i) Is TextBox Then
            If TypeOf .Controls(i) Is TextBox And .Controls(i).Enabled Then ' Hugo 06/01/2003 colocada para não pegar os campos desabilitados
'                If Trim(.Controls(i).Text) <> "" Then
                If (Trim(.Controls(i).Text) <> "") And (Len(Trim(.Controls(i).Text)) < 4000) Then
                    If strWhere <> "" Or InStr(1, strCondicao, "AND") > 0 Or InStr(1, strCondicao, "WHERE") > 0 Then
                        strCondicao = strCondicao & " AND "
                    Else
                        strCondicao = " WHERE "
                    End If
                    If UCase(Mid(.Controls(i).Name, 4, 3)) = "DTM" Then
                        If gblnDataValida(.Controls(i), True) Then
'                            strCondicao = strCondicao & strApelido & "." & strNomeCampo & " = '" & Format(.Controls(i).Text, "YYYY/MM/DD") & "'"
                            If (bytDBType = EDatabases.SQLServer) Then
                                strCondicao = strCondicao & strApelido & "." & strNomeCampo & " = '" & Format(.Controls(i).Text, "YYYY/MM/DD") & "'"
                            ElseIf (bytDBType = EDatabases.Oracle) Then
                                strCondicao = strCondicao & strApelido & "." & strNomeCampo & " = " & gstrFormataDataOracle(.Controls(i).Text, "YYYY/MM/DD")
                            End If
                        Else
                            Exit Function
                        End If
                    
                    ElseIf UCase(Mid(.Controls(i).Name, 4, 3)) = "DBL" Then 'Nino
                        'If (bytDBType = EDatabases.Oracle) Then
                            'strCondicao = strCondicao & strApelido & "." & strNomeCampo & " LIKE '" & Val(.Controls(i).Text) & "%'"
                            'strCondicao = strCondicao & gstrCONVERT(CDT_NVARCHAR, strApelido & "." & strNomeCampo) & " LIKE '" & gstrConvVrParaSql(.Controls(i).Text) & "%'"
                            strCondicao = strCondicao & strApelido & "." & strNomeCampo & " = " & gstrConvVrParaSql(.Controls(i).Text)
                        'End If
                    
                    ElseIf InStr(UCase(.Controls(i).Name), "INTCEP") > 0 Then
                        strCondicao = strCondicao & strApelido & "." & strNomeCampo & " LIKE '" & gstrValorSemMascara(.Controls(i).Text) & "%'"
                    ElseIf InStr(UCase(.Controls(i).Name), "CNPJ") > 0 Or InStr(UCase(.Controls(i).Name), "CPF") > 0 Then
                        strCondicao = strCondicao & strApelido & "." & strNomeCampo & " LIKE '" & gstrValorSemMascara(.Controls(i).Text) & "%'"
                    Else
                        'Nao entendi porque que é convertido como valor, e estava prejudicando consulta de campos string quando possui ponto
                        'strCondicao = strCondicao & " RTRIM(LTRIM(UPPER(" & strApelido & "." & strNomeCampo & "))) LIKE '" & Trim(UCase(gstrConvVrParaSql(.Controls(i).Text))) & "%'"
                        strCondicao = strCondicao & " RTRIM(LTRIM(UPPER(" & strApelido & "." & strNomeCampo & "))) LIKE '" & Trim(UCase(.Controls(i).Text)) & "%'"
                    End If
                End If
            ElseIf TypeOf .Controls(i) Is DataCombo Then
                
                If Trim(.Controls(i).Text) <> "" Then
                    If Not .Controls(i).MatchedWithList Then
'###############################################################
                        If .Controls(i).Tag <> "" And UCase$(Left(.Controls(i).Name, 3)) = "INT" Then
                            strSQLNotMatched = Left(.Controls(i).Tag, InStr(.Controls(i).Tag, ",") - 1)
                                        
                            If Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), " ")) = "" Then
                                strSQLNotMatched = strSQLNotMatched & " FROM " & Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), Len(.Controls(i).Tag))
                            Else
                                strFinalString = Right(Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), " ") + 5), 6)
                                Select Case Trim(UCase(strFinalString))
                                    Case "ORDER", "GROUP", "WHERE"
                                        strSQLNotMatched = strSQLNotMatched & " FROM " & Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), " "))
                                    Case Else
                                        If Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), ",")) <> "" Then
                                            strSQLNotMatched = strSQLNotMatched & " FROM " & Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl")), ","))
                                        ElseIf Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), ";")) <> "" Then
                                            strSQLNotMatched = strSQLNotMatched & " FROM " & Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), ";"))
                                        Else
                                            strSQLNotMatched = strSQLNotMatched & " FROM " & Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), Len(.Controls(i).Tag))
                                        End If
                                End Select
                            End If
                            strSQLNotMatched = strSQLNotMatched & " WHERE " & Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStrRev(.Controls(i).Tag, ";"))
                            strSQLNotMatched = strSQLNotMatched & " ='" & .Controls(i).Text & "'"
                
                            Set gobjBanco = New clsBanco
                                        
                            If gobjBanco.CriaADO(strSQLNotMatched, 5, adoResultado) Then
                                If adoResultado.EOF Then
                                                
                                    On Error Resume Next
                                    strLabel = .Controls(Replace(.Controls(i).Name, "dbc", "lbl")).Caption & " "
                                    On Error GoTo 0
                                                
                                    ExibeMensagem "O valor digitado para o campo " & strLabel & "não está correto!"
    
                                    .Controls(i).SetFocus
                                    Screen.MousePointer = vbNormal
                                    Exit Function
                                Else
                                    .Controls(i).BoundText = adoResultado(0) '"[" & Trim(.Controls(i).Text) & "]"
                                End If
                            End If
                                        
                            Set gobjBanco = Nothing
                        Else
                            On Error Resume Next
                            strLabel = .Controls(Replace(.Controls(i).Name, "dbc", "lbl")).Caption & " "
                            On Error GoTo 0
                                                    
                            ExibeMensagem "O valor digitado para o campo " & strLabel & "não está correto!"
        
                            If .Controls(i).Enabled Then
                                .Controls(i).SetFocus
                            End If
                            Screen.MousePointer = vbNormal
                            Exit Function
                        End If
                        
'###############################################################
                        strTextTemp = .Controls(i).Text
                        frmForm.MantemForm gstrPreencherLista
                        .Controls(i).Text = strTextTemp
                    End If
                    If strWhere <> "" Or InStr(1, strCondicao, "AND") > 0 Or InStr(1, strCondicao, "WHERE") > 0 Then
                        strCondicao = strCondicao & " AND "
                    Else
                        strCondicao = " WHERE "
                    End If
                    strCondicao = strCondicao & strApelido & "." & strNomeCampo & " = '" & gstrConvVrParaSql(.Controls(i).BoundText) & "'"
                End If
            End If
        End If
    Next i
End With

intPos = InStr(1, UCase(strSql), "SELECT") + InStr(1, UCase(strSql), "SP_")

If intPos = 0 Then
    strSql = "SELECT * FROM " & strTabela
End If

LocalizarGeral = strSql & " " & strWhere & " " & strCondicao & " " & strGroup & " " & strOrdena

Screen.MousePointer = vbNormal

End Function

Private Sub MudaGradeDBGrid(frmForm As Form)
gblnListViewComGrade = Abs(gblnListViewComGrade) - 1
VirificaGradeListView frmForm
GravaUsuario
End Sub

Public Function gstrDiaDaSemana(strData As String) As String
    '--------------------------------------------------------------'
    ' FUNÇÃO USADA PARA RETORNAR O DIA DA SEMANA.                  '
    '--------------------------------------------------------------'
    ' PARÂMETRO:                                                   '
    '                                                              '
    ' 1 - strData(Data utilizada para calcular o dia - Tipo String)'
    '                                                              '
    ' Obs:. Data no formato DD/MM/AA ou DD/MM/AAAA                 '
    '--------------------------------------------------------------'
    
    Dim vntDiaDaSemana As Variant
    If IsDate(strData) Then
        vntDiaDaSemana = Array("Domingo", _
                               "Segunda-Feira", _
                               "Terça-Feira", _
                               "Quarta-Feira", _
                               "Quinta-Feira", _
                               "Sexta-Feira", _
                               "Sábado")
        gstrDiaDaSemana = vntDiaDaSemana(Weekday(strData) - 1)
    End If
End Function

Public Sub gblnFilraCampos(ByRef tdb_Grid As TDBGrid)
Dim adoRec As ADODB.Recordset
Dim Col    As TrueOleDBGrid70.Column
Dim c      As Integer
Dim tmp    As String
Dim n      As Integer
Dim strAux As String

On Error GoTo err_gblnFilraCampos

    Set adoRec = New ADODB.Recordset
    Set adoRec = tdb_Grid.DataSource

    If adoRec.BOF And adoRec.EOF Then
        Exit Sub
    End If
    
    c = tdb_Grid.Col
    tdb_Grid.HoldFields
    tmp = ""
    
    For Each Col In tdb_Grid.Columns
    
        If Trim(Col.FilterText) <> "" And Trim(Col.FilterText) <> "%" Then

            n = n + 1

            If tmp <> "" Then
                tmp = tmp & " AND "
            End If
    
            If (adoRec.Fields(Col.DataField).Type = 3 Or adoRec.Fields(Col.DataField).Type = 131 Or adoRec.Fields(Col.DataField).Type = 139) And IsNumeric(Col.FilterText) Then
                tmp = tmp & Col.DataField & " = " & gstrConvVrParaSql(Col.FilterText)
                
            ElseIf adoRec.Fields(Col.DataField).Type = 135 Then
                If gblnDataValida(Col.FilterText) And Len(Col.FilterText) > 7 Then
                    'tmp = tmp & Col.DataField & " = " & gstrConvDtParaSql(Col.FilterText)
                    tmp = tmp & Col.DataField & " = " & Col.FilterText
                End If
                
            ElseIf Not adoRec.Fields(Col.DataField).Type = 131 Then
        
                If InStr(1, UCase(Col.DataField), "CPF") > 0 Or InStr(1, UCase(Col.DataField), "CNPJ") > 0 Then
                    strAux = gstrValorSemMascara(Col.Text)
                Else
                    strAux = Col.FilterText
                End If
                tmp = tmp & Col.DataField & " LIKE '" & strAux & "*'"
                
            Else
                Col.FilterText = ""
            End If
        
        End If
        
    Next
    
    If tmp <> "" Then
        
        adoRec.Filter = tmp
        
        If adoRec.EOF And adoRec.BOF Then
            adoRec.Filter = adFilterNone
            'adoRec.MoveLast
        End If
        
    Else
        adoRec.Filter = adFilterNone
    End If
    
    tdb_Grid.Col = c
    tdb_Grid.EditActive = True
    tdb_Grid.CurrentCellModified = True

err_gblnFilraCampos:

End Sub

Public Sub ChamaFormCadastro(frmForm As Form, _
                             objObjeto As Object, _
                    Optional strQuery As String)

    '-----------------------------------------------------------------'
    ' SUB USADA PARA CHAMAR O FORMULÁRIO (FORM) DE CADASTRO QUANDO    '
    ' CHAMADO DE OUTRA TELA (FORA DO MENU) E RETORNAR DADOS DO        '
    ' FORMULÁRIO DE CADASTRO PARA O DE ORIGEM                         '
    '-----------------------------------------------------------------'
    ' PARÂMETROS:                                                     '
    '                                                                 '
    ' 1 - frmForm (Formulário a ser carregado (chamado))              '
    ' 2 - objObjeto (objeto (lista) onde retornarão os dados 'ComBox, '
    '                ListView Etc.                                    '
    '-----------------------------------------------------------------'
    gstrQueryParamGeral = strQuery
    Set gobjGeral = objObjeto
    TrocaInconiDoObj frmForm, 3
    If frmForm.WindowState <> vbMaximized Then
        frmForm.Left = 0
        frmForm.Top = 0
        frmForm.BorderStyle = vbFixedDialog
    End If
    frmForm.Show
    frmForm.SetFocus
End Sub

Public Sub TrocaInconiDoObj(objObjeto As Object, _
                            bytTipo As Byte)
    Dim objControle As Object
    On Error Resume Next
    With frmIconiEspecial
        Select Case bytTipo
        Case vbBeginDrag
            objObjeto.DragIcon = .lbl_Drag.DragIcon
        Case vbEndDrag, vbCancel
            objObjeto.DragIcon = .lbl_Drop.DragIcon
        Case Else
            For Each objControle In objObjeto.Controls
                If LCase(Mid(objControle.Name, 1, 4)) = "cmd_" Then
                    objControle.MouseIcon = .lbl_Mao.MouseIcon
                    objControle.MousePointer = vbCustom
'--------- Verifica permisões do usuário e habilita ou desabilita o botão
                    If Val(objControle.Tag) <> 0 Then
                        If gstrPermissao(Val(objControle.Tag)) <> "" Then
                            If InStr(1, gstrPermissao(Val(objControle.Tag)), "2") = 0 Then
                                'objControle.Enabled = False
                            End If
                        End If
                    End If
                End If
            Next
            objObjeto.Icon = MDIMenu.Icon
        End Select
        Unload frmIconiEspecial
    End With
End Sub

Public Sub VerificaParametroCombox(objGeral As Object, _
                          Optional tlbBarraFermta As Toolbar)
    Dim btnBotao As Button
    If gobjGeral Is Nothing = False Then
        If tlbBarraFermta Is Nothing = False Then
            For Each btnBotao In tlbBarraFermta.Buttons
                If UCase(btnBotao.Key) = UCase(gstrAplicar) Then
                    btnBotao.Visible = True
                    Exit For
                End If
            Next
        End If
        Set objGeral = gobjGeral
        gblnRetornaRegistro = True
    End If
    Set gobjGeral = Nothing
End Sub

Public Sub VerificaObjParaAplicar(objGeral As Object, _
                         Optional strQuery As String)
    '--------------------------------------------------------------
    ' FUNÇÃO USADA PARA VEIRICAR SE UM OBJETO FOI INFORMADO PARA
    ' RETORNAR COM OS DADOS PARA OUTRO FORMULÁRIO E HABILITAR O
    ' BOTÃO APLICAR DA BARRA DE FERRAMENTAS
    '--------------------------------------------------------------
    ' PARÂMETROS:
    '
    ' 1 - objGeral (objeto do formulário que receberá os dados
    ' 2 - tlbBarraFermta (barra de ferramenta com o botão 'aplicar
    '--------------------------------------------------------------
    Set objGeral = Nothing
    If gobjGeral Is Nothing = False Then
        Set objGeral = gobjGeral
        gblnRetornaRegistro = True
        strQuery = gstrQueryParamGeral
    End If
    DoEvents
    Set gobjGeral = Nothing
End Sub

Public Function gblnVerificaPermissoes(pRotina As Integer, _
                                       strBandeira As String, _
                              Optional blnMenu As Boolean) As Boolean
    '---------------------------------------------------------------'
    ' FUNÇÃO USADA PARA VERIFICAR PERMISSÕES DE ACESSO DA ROTINA.   '
    '---------------------------------------------------------------'
    ' PARÂMETRO:                                                    '
    '                                                               '
    ' 1 - pRotina(Rotina - Tipo Long)                             '
    '---------------------------------------------------------------'
    Dim adoPermissoes   As ADODB.Recordset
    Dim adoResultado    As ADODB.Recordset
    Dim strPermissao    As String
    Dim strSql          As String
    Dim vntBotao        As Variant
    Dim i               As Integer
    
    On Error GoTo err_VerificaPermissoes
    
    gblnVerificaPermissoes = True
    
    If pRotina = 0 Then Exit Function
    
    strPermissao = gstrPermissao(pRotina)
    
    If strPermissao = "*" Then Exit Function
    
        If blnMenu Then
            With MDIMenu.actBarra.Bands(strBandeira)
                If Trim(strPermissao) = "" Then
                    gblnVerificaPermissoes = False
                    Exit Function
                ElseIf InStr(1, strPermissao, "2") = 0 Then
                    gblnVerificaPermissoes = False
                    Exit Function
                End If
            End With
        Else
            With MDIMenu.actBarra.Bands(strBandeira)
                For i = 1 To 5
                    Select Case Mid(strPermissao, i, 1)
                    Case "1"
                        .Tools(i - 1).Enabled = False
                    Case "2"
                        '.Tools(i - 1).Enabled = True
                    End Select
                Next i
            End With
            
            'Botões Específicos
            strSql = ""
            strSql = strSql & "SELECT * FROM "
            strSql = strSql & gstrItemPermissaoEspecifica & " "
            strSql = strSql & "WHERE intItem = " & pRotina & " "
            strSql = strSql & "ORDER BY intPosicao"
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                With adoResultado
                    If Not (.BOF And .EOF) Then
                        For i = 6 To 6 + .RecordCount - 1
                            With MDIMenu.actBarra.Bands(gstrBtnArquivo)
                                Select Case Mid(strPermissao, i, 1)
                                Case "1"
                                    .Tools(CStr(Trim(adoResultado!strMenu))).Enabled = False
                                Case "2"
                                    '.Tools(i + 1).Enabled = True
                                End Select
                            End With
                            .MoveNext
                        Next
                    End If
                End With
            End If
        End If

    Exit Function
err_VerificaPermissoes:
End Function

Public Function VerificaPermissoesNaString(pRotina As Integer, intPosicao As Integer) As Boolean
Dim strPermissao    As String
Dim strSql          As String
    
    On Error GoTo err_VerificaPermissoesNaString
    
    VerificaPermissoesNaString = True
    
    If pRotina = 0 Then Exit Function
    
    strPermissao = gstrPermissao(pRotina)
    
    If strPermissao = "*" Then Exit Function
    
    If Trim(strPermissao) = "" Then
        VerificaPermissoesNaString = False
        Exit Function
    ElseIf InStr(1, strPermissao, "2") = 0 Then
        VerificaPermissoesNaString = False
        Exit Function
    End If
        
    If Mid(strPermissao, intPosicao, 1) = "1" Then
        VerificaPermissoesNaString = False
    End If
        
    Exit Function

err_VerificaPermissoesNaString:

End Function

Public Function gstrPermissao(mintCodigo) As String
    '---------------------------------------------------------------------
    ' FUNÇÃO USADA PARA PROCURAR, NO VETOR DE SEGURANÇA,
    ' A PERMISSÃO CORRESPONDENTE AO CÓDIGO INFORMADO
    '---------------------------------------------------------------------
    ' PARÂMETROS:
    '
    ' 1 - mintCodigo(Código de segurança correspondente ao
    '     formulário/relatório ativado)
    '---------------------------------------------------------------------
    Dim intLoop As Integer
    On Error GoTo err_gstrPermissao
    gstrPermissao = ""
    For intLoop = 0 To UBound(vetPermissoes)
        gstrPermissao = "*"
        If vetPermissoes(intLoop).intCodigo = mintCodigo Then
            gstrPermissao = Trim(vetPermissoes(intLoop).strPermissao)
            Exit Function
        End If
    Next intLoop
err_gstrPermissao:
End Function

Sub SetBotaoBarraFerramenta(cmlbtnBotao As MSComctlLib.Button, _
                            tlbBarraFermta As Toolbar, _
                   Optional intCaracter As Integer, _
                   Optional strBotao As String)
    
    '---------------------------------------------------------------------
    ' SUB USADA PARA PROCURAR, NA BARRA DE FERRAMENTAS, O BOTÃO
    ' CORRESPONDENTE À TECLA PRESSIONADA
    '---------------------------------------------------------------------
    ' PARÂMETROS:
    '
    ' 1 - cmlbtnBotao(Para retornar o botão correspondente, se encontrado
    ' 2 - tlbBarraFermta(A barra de ferramentas onde seré feita a procura
    ' 3 - intCaracter(indica a tecla pressionada
    ' 4 - intCaracter(Caracter digitado - Tipo Integer)
    ' 4 - strBotao(Para procurar por um botão específico
    '---------------------------------------------------------------------
                   
    For Each cmlbtnBotao In tlbBarraFermta.Buttons
        If (UCase(cmlbtnBotao.Key) = UCase(gstrFechar) And intCaracter = vbKeyEscape) _
           Or (UCase(cmlbtnBotao.Key) = UCase(gstrSalvar) And ((intCaracter = vbKeyReturn And gblnTeclaEnterIgualTab = False) Or intCaracter = 19)) _
           Or (UCase(cmlbtnBotao.Key) = UCase(gstrImprimir) And intCaracter = 9) _
           Or (UCase(cmlbtnBotao.Key) = UCase(gstrNovo) And intCaracter = 14) _
           Or ((Trim(strBotao) <> "") And UCase(cmlbtnBotao.Key) = UCase(strBotao)) Then
            Exit For
        End If
    Next
End Sub


Public Function gstrMsgErroADO(Err, strQueryAux As String) As String
    Dim strAux          As String
    Dim intInd          As Integer
    Dim intAux          As Integer
    Dim blnConcatena    As Boolean
    
    If InStr(UCase(strQueryAux), "DELETE") Then
        gstrMsgErroADO = gstrMsgExclusao(Err, strQueryAux)
    ElseIf Err.Number = -2147467259 Then
        For intInd = 1 To Len(Trim(Err.Description))
            If Mid(Err.Description, intInd, 1) = "." Then
                blnConcatena = True
            End If
            If blnConcatena Then
                Select Case UCase(Mid(Err.Description, intInd, 1))
                Case Chr(vbKeyA) To Chr(vbKeyZ)
                    strAux = strAux & Mid(Err.Description, intInd, 1)
                Case "'"
                    Exit For
                End Select
            End If
        Next
        gstrMsgErroADO = "O campo " & strAux & " tem que ser informado."
    ElseIf Err.Number = -2147217900 Then
        If InStr(Err.Description, "IX_") > 0 Then
            intAux = InStr(Err.Description, "IX_")
            For intInd = (intAux + 3) To Len(Trim(Err.Description))
                If Mid(Err.Description, intInd, 1) = "'" Then
                    Exit For
                ElseIf Mid(Err.Description, intInd, 1) = "_" Then
                    strAux = strAux & Chr(vbKeySpace)
                Else
                    strAux = strAux & Mid(Err.Description, intInd, 1)
                End If
            Next
            gstrMsgErroADO = "Não é permitido dois registros com o mesmo valor para a coluna '" & strAux & "'"
        ElseIf InStr(Err.Description, "FK_") > 0 Then
            intAux = InStr(Err.Description, "FK_")
            For intInd = (intAux + 3) To Len(Trim(Err.Description))
                If Mid(Err.Description, intInd, 1) = "'" Or Mid(Err.Description, intInd, 1) = "_" Then
                    Exit For
                ElseIf Mid(Err.Description, intInd, 1) = "_" Then
                    strAux = strAux & Chr(vbKeySpace)
                Else
                    strAux = strAux & Mid(Err.Description, intInd, 1)
                End If
            Next
            gstrMsgErroADO = "O campo '" & strAux & "' tem que ser informado"
        ElseIf InStr(Err.Description, "NULL") > 0 Then
            intAux = InStr(Err.Description, "'") + 1
            For intInd = (intAux + 3) To Len(Trim(Err.Description))
                If Mid(Err.Description, intInd, 1) = "'" Or Mid(Err.Description, intInd, 1) = "_" Then
                    Exit For
                ElseIf Mid(Err.Description, intInd, 1) = "_" Then
                    strAux = strAux & Chr(vbKeySpace)
                Else
                    strAux = strAux & Mid(Err.Description, intInd, 1)
                End If
            Next
            gstrMsgErroADO = "O campo '" & strAux & "' tem que ser informado"
        Else
            gstrMsgErroADO = ""
        End If
    Else
        gstrMsgErroADO = ""
    End If
End Function

Public Function gstrMsgExclusao(Err, strQueryAux As String) As String

'******************************************************************************************
' Data: 15/04/2003
' Alteração: - Adicionada consulta às tabelas de catálogo do Oracle a fim de trazer o nome
'            da tabela referenciada pela referência que causou o erro de deleção.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strAux          As String
    Dim strMsg          As String
    Dim intAux          As Long
    Dim strTabela       As String
    
    Dim adoResultado    As New ADODB.Recordset
    Dim strSql          As String
            
    If (bytDBType = EDatabases.SQLServer) Then
        If InStr(UCase(Err.Description), "TABLE") Then
            intAux = InStr(UCase(Err.Description), "TABLE") + 7
        ElseIf InStr(UCase(Err.Description), "TBL") Then
            intAux = InStr(UCase(Err.Description), "TBL")
        End If
    
    ElseIf (bytDBType = EDatabases.Oracle) Then
        If InStr(UCase(Err.Description), "CPDMASTER.") Then
            intAux = InStr(UCase(Err.Description), "CPDMASTER.") + 10
        End If
    
    End If
    
    If intAux > 0 Then
        Do
            strAux = strAux & Mid(Err.Description, intAux, 1)
            intAux = intAux + 1
'        Loop Until Mid(Err.Description, intAux, 1) = "'"
        Loop Until (Mid(Err.Description, intAux, 1) = "'" And bytDBType = EDatabases.SQLServer) Or _
                    (Mid(Err.Description, intAux, 1) = ")" And bytDBType = EDatabases.Oracle) Or _
                    (intAux > Len(Err.Description))
        
        If (bytDBType = EDatabases.Oracle) Then
            
            strSql = "SELECT TABLE_NAME "
            strSql = strSql & "From ALL_CONSTRAINTS "
            strSql = strSql & "Where "
            strSql = strSql & "OWNER = 'CPDMASTER' AND "
            strSql = strSql & "CONSTRAINT_TYPE = 'R' AND "
            strSql = strSql & "CONSTRAINT_NAME = '" & strAux & "' "
    
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                If Not adoResultado.EOF Then
                    strAux = adoResultado("TABLE_NAME")
                End If
            End If
            
        End If
        
        strTabela = gstrCatalogoTabelaGeral(strAux)
        strMsg = strMsg & "Este registro não pode ser excluído porque ele " & Chr(10) & "está sendo utilizando em " & strTabela
        gstrMsgExclusao = strMsg
    
    Else
        gstrMsgExclusao = ""
    
    End If

End Function

Public Function gstrCatalogoTabelaGeral(strTabela As String) As String
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    strSql = strSql & "SELECT strDescricao FROM " & gstrCatalogoTabela & " "
    strSql = strSql & "WHERE UPPER(strTabela) = '" & UCase(Trim(strTabela)) & "'"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            gstrCatalogoTabelaGeral = adoResultado!strDescricao
        Else
            gstrCatalogoTabelaGeral = strTabela
        End If
    Else
        gstrCatalogoTabelaGeral = strTabela
    End If
End Function


Public Function gstrQueryUF(Optional blnRetornaDescricao As Boolean) As String
    
    '--------------------------------------------------------------------------------
    ' FUNCAO USADA PARA MONTAR UMA QUERY ESPECÍFICA PARA
    ' SELECIONAR AS UNIDADES DA FEDERAÇÃO
    '--------------------------------------------------------------------------------
    ' PARÂMETROS:
    ' 1 - blnParaListView(flag indicando se o objeto a ser preenchido é um ListView)
    '--------------------------------------------------------------------------------
    
    Dim strSql  As String
    If blnRetornaDescricao Then
        strSql = strSql & "SELECT PKId, strEstado, strSigla "
        strSql = strSql & "FROM " & gstrUF & " ORDER BY strEstado"
    Else
        strSql = strSql & "SELECT PKId, strSigla "
        strSql = strSql & "FROM " & gstrUF & " ORDER BY strSigla"
    End If
    gstrQueryUF = strSql
End Function

Public Function gblnCepValido(objCep As Object, Optional objLogradouro As Object, Optional objMunicipio As Object) As Boolean
    '---------------------------------------------------------------'
    ' FUNÇÃO USADA PARA VERIFICAR UM CEP DE ACORDO COM O LOGRADOURO '
    ' OU O MUNICÍPIO                                                '
    '---------------------------------------------------------------'
    ' PARÂMETRO:                                                    '
    '                                                               '
    ' 1 - objCep - Cep a ser verificado.(Tipo Object)               '
    ' 1 - objLogradouro - (Tipo Object)                             '
    ' 1 - objMunicipio -  (Tipo Object)                             '
    '---------------------------------------------------------------'
    
    Dim adoResultado As ADODB.Recordset
    Dim strSql       As String
    
    'RETIRAR ISSO QUANDO FOR FAZER A NOVA ROTINA DE VALIDAÇÃO
    gblnCepValido = True: Exit Function
    
    If objCep = "" Then
        Exit Function
    ElseIf objLogradouro Is Nothing And objMunicipio Is Nothing Then
        ExibeMensagem "O logradouro e/ou município tem que ser informado."
        Exit Function
    End If
    
    If objLogradouro Is Nothing Then
VerificacaoPorMunicipio:
        If objMunicipio = "" Then
            Exit Function
        End If

        strSql = ""
        strSql = strSql & "Select intCepInicial, intCepFinal From " & gstrCidade & " "
        strSql = strSql & "Where PKId = " & objMunicipio.BoundText
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                If gstrValorSemMascara(objCep) < adoResultado!intCepInicial Or gstrValorSemMascara(objCep) > adoResultado!intCepFinal Then
'                    ExibeMensagem "Cep inválido para o município selecionado."
                    objCep.SetFocus
                    Exit Function
                End If
            End If
        End If
    
    ElseIf objMunicipio Is Nothing Then
VerificacaoPorLogradouro:
            If objLogradouro = "" Then
                Exit Function
            End If

            strSql = ""
            strSql = strSql & "Select intCep From " & gstrCepsLogradouro & " "
            strSql = strSql & "Where intCep = " & gstrValorSemMascara(objCep) & " "
            strSql = strSql & "And intLogradouro = " & objLogradouro.BoundText
            
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                If adoResultado.EOF Then
'                    ExibeMensagem "Cep não cadastrado para o logradouro selecionado."
                    objCep.SetFocus
                    Exit Function
                End If
            End If
        
    Else
        If gintMunicipioEmpresa = objMunicipio.BoundText Then
            GoTo VerificacaoPorLogradouro
        Else
            GoTo VerificacaoPorMunicipio
        End If
    End If
    gblnCepValido = True
End Function

Function gblnExisteValorNaTabela(strTabela As String, _
                                 strCampo As String, _
                                 vntValor As Variant, _
                        Optional blnAlterando As Boolean, _
                        Optional objLista As ListView) As Boolean
    
    Dim adoTemp  As ADODB.Recordset
    Dim strSql   As String
    Dim gstrDado As String
    
    gblnExisteValorNaTabela = True
    
    strSql = ""
    strSql = strSql & "Select Max(" & strCampo & ") "
    strSql = strSql & "From " & strTabela
    strSql = strSql & " Where PKId = 1000"
    
    Set gobjBanco = New clsBanco
    If Not gobjBanco.CriaADO(strSql, 5, adoTemp) Then
        Exit Function
    End If
    
    Select Case adoTemp.Fields(0).Type
    Case adChar, adVarChar, adVarWChar, adLongVarChar
        If Len(vntValor) = 0 Then
            gstrDado = "Null"
        Else
            gstrDado = "'" & vntValor & "'"
        End If
    Case adNumeric, adCurrency
        gstrDado = gstrConvVrParaSql(vntValor)
    Case adDate, adDBTimeStamp
        gstrDado = gstrConvDtParaSql(vntValor)
    Case adBoolean
        gstrDado = Abs(vntValor)
    Case adInteger
        gstrDado = Val(vntValor)
    Case Else
        gstrDado = vntValor
    End Select
    
    strSql = ""
    strSql = strSql & "Select " & strCampo & " "
    strSql = strSql & "From " & strTabela & " "
    strSql = strSql & "Where " & strCampo & " = " & gstrDado
    
    If blnAlterando Then
        strSql = strSql & " And PKId <> " & objLista.SelectedItem.Tag
    End If
    
    If gobjBanco.CriaADO(strSql, 5, adoTemp) Then
        If Not adoTemp.EOF Then
            Exit Function
        End If
    End If
    
    gblnExisteValorNaTabela = False
End Function


Public Function gstrConvDtParaSql(vntData As Variant, _
                         Optional blnParaSelect As Boolean) As String
    
    '-------------------------------------------------------------
    ' FUNÇÃO USADA PARA CONVERTER A DATA 'vntData' PARA O FORMATO
    ' 'YYYY/MM/DD' OR RETORNAR NULO PARA O SQL
    '-------------------------------------------------------------
    ' PARAMETRO:
    ' 1 - vntData (Data a ser convertida)
    '----------------------------------------------------------
    

'******************************************************************************************
' Data: 11/03/2003
' Alteração: - Todo valor que for passado para um campo date no Oracle deve ser traduzido
'            antes pela função TO_DATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    If gblnDataValida(vntData) Then
'        If blnParaSelect Then
'            gstrConvDtParaSql = "#" & Format(vntData, "mm/dd/yyyy hh:mm:ss") & "#"
'        Else
'            gstrConvDtParaSql = "'" & Format(vntData, "yyyy/mm/dd hh:mm:ss") & "'"
'        End If
        Select Case bytDBType
            Case EDatabases.SQLServer
                If blnParaSelect Then
                    gstrConvDtParaSql = "#" & Format(vntData, "mm/dd/yyyy hh:mm:ss") & "#"
                Else
                    gstrConvDtParaSql = "'" & Format(vntData, "yyyy/mm/dd hh:mm:ss") & "'"
                End If
            
            Case EDatabases.Oracle
                gstrConvDtParaSql = gstrFormataDataOracle(CStr(vntData))
        
        End Select
    Else
        gstrConvDtParaSql = "NULL"
    End If
End Function


Public Function glngPegaProximaChave(strTabela As String, _
                                     strCampo As String, _
                            Optional strCampoCond1 As String, _
                            Optional vntValor1 As Variant, _
                            Optional strCampoCond2 As String, _
                            Optional vntValor2 As Variant) As Long
    '--------------------------------------------------------------'
    ' FUNÇÃO USADA PARA RETORNAR A ÚLTIMA CHAVE DA TABELA.         '
    '--------------------------------------------------------------'
    ' PARÂMETROS:                                                  '
    '                                                              '
    ' 1 - sTAbela(Tabela - Tipo String)                            '
    ' 2 - sCampo(Campo chave da tabela - Tipo String)              '
    '--------------------------------------------------------------'
    
'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 25/03/2003
' Alteração: - Substituição da variável strISNULL pela função gstrISNULL. Ver definição da
'            função para maiores esclarecimentos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim adoResultado  As ADODB.Recordset
    Dim strSql        As String
    Dim mobjAuxBanco  As Object

    strSql = ""
'    strSql = strSql & "SELECT ISNULL(MAX(" & Trim(strCampo) & "), 0) Maximo "
'    strSQL = strSQL & "SELECT " & strISNULL & "(MAX(" & Trim(strCampo) & "), 0) Maximo "
    strSql = strSql & "SELECT " & gstrISNULL("MAX(" & Trim(strCampo) & ")", "0") & " Maximo "
    strSql = strSql & "FROM " & strTabela
    If Trim(strCampoCond1) <> "" Then
        strSql = strSql & " WHERE " & strCampoCond1 & " = " & vntValor1
    End If
    If Trim(strCampoCond2) <> "" Then
        strSql = strSql & " AND " & strCampoCond2 & " = " & vntValor2
    End If
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        glngPegaProximaChave = adoResultado!Maximo + 1
        Set gobjBanco = Nothing
        adoResultado.Close
        Set adoResultado = Nothing
    End If
End Function



Public Sub OrdenaColunaClicada(lvwLisa As ListView, _
                               ColumnHeader As MSComctlLib.ColumnHeader)
    '---------------------------------------------------------------'
    ' SUB USADA PARA ORDENAR O LISTVIEW PELA COLUNA CLICADA.        '
    '---------------------------------------------------------------'
    ' PARÂMETROS:                                                   '
    '                                                               '
    ' 1 - lvw(ListView - Tipo ListView)                             '
    ' 2 - ColumnHeader(Coluna Clicada)                              '
    '---------------------------------------------------------------'
    Dim intInd  As Integer
    intInd = lvwLisa.SortKey
    If ColumnHeader.Index - 1 <> intInd Then
        lvwLisa.SortKey = ColumnHeader.Index - 1
    Else
        If lvwLisa.SortOrder = lvwAscending Then
            lvwLisa.SortOrder = lvwDescending
        Else
            lvwLisa.SortOrder = lvwAscending
        End If
    End If
    lvwLisa.Sorted = True
End Sub


Public Sub ComboGeral(frmForm As Form, _
                      objGeral As Object, _
                      lvwLista As ListView, _
                      strTabela As String, _
             Optional strPKId As String, _
             Optional strQuery As String)
      
    '--------------------------------------------------------------'
    ' SUB USADA PARA AUTOMATIZAR O PREENCHIMENTO DE LISTAGEM NO    '
    ' FORMULÁRIO DE ORIGEM (CHAMADOR)
    '--------------------------------------------------------------'
    ' PARÂMETROS:                                                  '
    '                                                              '
    ' 1 - frmForm(Formulário de origem)                            '
    ' 2 - objGeral(Objeto que será preenchido)                     '
    ' 3 - lvwLista(lista que contem o dado escolhido)              '
    ' 3 - strTabela(Tabela de onde será lido os dados)             '
    ' 4 - strPKId(Chave da tabela)                                 '
    '--------------------------------------------------------------'
    On Error GoTo ErroObjeto
    If objGeral Is Nothing = False Then
        With lvwLista
            If .SelectedItem.Selected Then
                LeDaTabelaParaObj strTabela, objGeral, strPKId, strQuery
                If TypeOf objGeral Is ComboBox Then
                    objGeral.ListIndex = gintIndiceCBO(objGeral, .SelectedItem.Tag)
                ElseIf TypeOf objGeral Is DataCombo Then
                    objGeral.BoundText = .SelectedItem.Tag
                ElseIf TypeOf objGeral Is ListView Then
                    Call gblnEncontroItemNoListView(objGeral, .SelectedItem.Tag, lvwTag)
                End If
                If objGeral.Enabled Then
                    objGeral.SetFocus
                End If
                Set objGeral = Nothing
            End If
        End With
    Else
        Exit Sub
    End If
    GoTo FimErro
    
ErroObjeto:
    Resume FimErro
    
FimErro:
    Unload frmForm
    
End Sub
Sub PesquisaListView(intKeyCode As Integer, _
                     objObjeto As Object, _
                     lvw_Lista As ListView, _
                     blnAlterando As Boolean, _
            Optional intLocalProcura As Integer, _
            Optional intTipoProcura As Integer)
    If objObjeto.Name = Screen.ActiveControl.Name Then
        If blnAlterando = False Then
            If gblnEncontroItemNoListView(lvw_Lista, objObjeto.Text, _
                                          intLocalProcura, intTipoProcura) Then
                If intKeyCode = vbKeyReturn Then
                    lvw_Lista.SetFocus
                    SendKeys Chr(vbKeySpace)
                    DoEvents
                End If
            End If
        End If
        If intKeyCode = vbKeyReturn Then
            If gblnTeclaEnterIgualTab Then
                EnviaTeclaTab vbKeyReturn
            End If
        End If
    End If
End Sub

Public Function gstrMOuF(Optional blnMasculino As Boolean, _
                         Optional blnPalavra As Boolean, _
                         Optional vntLetra As Variant)
                         
    '-------------------------------------------------------------------------------
    ' FUNÇÃO USADA PARA RETORNAR AS PALAVRAS 'Masculino ou Feminino'
    ' OU AS LETRAS 'F ou M' DE ACORDO COM OS FLAGS INFORMADOS
    '-------------------------------------------------------------------------------
    ' PARÂMETRO:
    '
    ' 1 - blnMasculino(Flag indicando se é feminino ou masculino)
    ' 2 - blnPalavra(Flag indicando se retorna letra 'M ou F' ou palavra
                     'Masculino ou Feminico')
    ' 3 - vntLetra (se o parâmetro letra for informado o flag - blnMasculino -
    '               retorna falso ou verdadeiro 'masculino ou feminino'
    '               conforme o valor informado - 'M ou outro valor')
    '------------------------------------------------------------------------------
                         
    If IsMissing(vntLetra) = False Then
        If UCase(vntLetra) = "M" Then
            blnMasculino = True
        End If
    End If
    If blnPalavra Then
        If blnMasculino Then
            gstrMOuF = "Masculino"
        Else
            gstrMOuF = "Feminino"
        End If
    Else
        If blnMasculino Then
            gstrMOuF = "M"
        Else
            gstrMOuF = "F"
        End If
    End If
End Function



'Public Function gstrRetornaSimOuNao(vntFlag As Variant) As String
'    If vntFlag = True Or Trim(vntFlag) = "True" Then
'        gstrRetornaSimOuNao = "Sim"
'    ElseIf vntFlag = False Or Trim(vntFlag) = "False" Then
'        gstrRetornaSimOuNao = "Não"
'    ElseIf Abs(vntFlag) = 0 Then
'        gstrRetornaSimOuNao = "Não"
'    ElseIf Abs(vntFlag) = 1 Then
'        gstrRetornaSimOuNao = "Sim"
'    End If
'End Function
'
Public Function gstrQueryLogradouro(Optional strTabela As String, _
                                    Optional strCampoCondicao As String, _
                                    Optional strChaveEstrangeira As String, _
                                    Optional blnExcluidos As Boolean = True) As String

    
'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 06/03/2003
' Alteração: - Adaptação dos outer joins.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 10/03/2003
' Alteração: - Alteração do comando SELECT devido a incompatibilidades de estrutura dos
'            outer joins entre o SQL Server e o Oracle. Os joins da cláusula FROM foram
'            substituídos por joins correspondentes na cláusula WHERE.
'            - O comando CONVERT fora retirado por ter-se verificado que não era necessário
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 25/03/2003
' Alteração: - Substituição da variável strISNULL pela função gstrISNULL. Ver definição da
'            função para maiores esclarecimentos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql  As String
    
    strSql = ""
    strSql = strSql & "SELECT L.PKId, "
    strSql = strSql & " RTRIM(LTRIM(L.strDescricao)) " & strCONCAT & gstrISNULL("TL.strSigla", "''", "', '") & strCONCAT & " RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''") & _
                        strCONCAT & gstrISNULL("U.strDescricao", "' '", "', '") & strCONCAT & gstrISNULL("U.strDescricao", "''") & ")) AS Logradouro "
    
    If strTabela <> "" And strTabela <> gstrBairro Then
        strSql = strSql & ", " & strTabela & ".intNumero "
        strSql = strSql & ", " & strTabela & ".strComplemento "
    End If
    
    strSql = strSql & "FROM " & gstrLogradouro & " L, "
    strSql = strSql & gstrTituloLogradouro & " U, "
    strSql = strSql & gstrTipoLogradouro & " TL "
    
    If strTabela <> "" Then
        strSql = strSql & ", " & strTabela
    End If
    
    strSql = strSql & " WHERE L.intTituloLogradouro " & strOUTJSQLServer & "= U.PKId " & strOUTJOracle
    
    If strChaveEstrangeira <> "" Then
        strSql = strSql & " AND L.PKId = " & strTabela & "." & strChaveEstrangeira
    End If
    
    strSql = strSql & " AND L.intTipoLogradouro " & strOUTJSQLServer & "= TL.PKId " & strOUTJOracle
    
    If strCampoCondicao <> "" Then
        strSql = strSql & " AND " & strTabela & "." & strCampoCondicao
    End If
    
    If Not blnExcluidos Then
        strSql = strSql & " AND L.Dtmdtexclusao Is Null"
    End If

    strSql = strSql & " ORDER BY L.strDescricao "
    
    gstrQueryLogradouro = strSql
        
End Function

Public Function gstrQueryTituloLogradouro(Optional blnRetornaDescricao As Boolean) As String
    
    '--------------------------------------------------------------------------------
    ' FUNCAO USADA PARA MONTAR UMA QUERY ESPECÍFICA PARA
    ' SELECIONAR O TITULO DO LOGRADOURO
    '--------------------------------------------------------------------------------
    ' PARÂMETROS:
    ' 1 - blnParaListView(flag indicando se o)
    '--------------------------------------------------------------------------------

    Dim strSql  As String
    If blnRetornaDescricao Then
        strSql = strSql & "SELECT PKId, strDescricao, strSigla "
        strSql = strSql & "FROM " & gstrTituloLogradouro & " ORDER BY strDescricao"
    Else
        strSql = strSql & "SELECT PKId, strSigla "
        strSql = strSql & "FROM " & gstrTituloLogradouro & " ORDER BY strSigla"
    End If
    gstrQueryTituloLogradouro = strSql
End Function

Public Function gstrQueryCidade()
    Dim strSql  As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao FROM " & gstrCidade & " "
    strSql = strSql & "ORDER BY strDescricao"
    gstrQueryCidade = strSql
End Function

Public Sub LeCodigoEspecifico(objCodigo As Object, _
                               strTabela As String, _
                               objLista As Object, _
                      Optional strCampo As String)
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    strSql = ""
    If Trim(strCampo) <> "" Then
        strSql = strSql & "SELECT " & Trim(strCampo) & " "
        strSql = strSql & "AS strCodigo FROM " & strTabela & " "
    Else
        strSql = strSql & "SELECT strCodigo FROM " & strTabela & " "
    End If
    strSql = strSql & "WHERE PKId = " & gstrItemData(objLista)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            objCodigo = adoResultado!strCodigo
        End If
    End If
End Sub

Public Sub LeImagemLogotipo(imgBrasao, _
                            imgLogotipo, _
                   Optional fldEmpresa As Object, _
                   Optional fldEstado As Object)

'******************************************************************************************
' Data: 06/03/2003
' Alteração: - Adaptação dos outer joins.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql      As String
    Dim adoEmpresa  As ADODB.Recordset
    If fldEmpresa Is Nothing = False Then
        
        On Error Resume Next
        
        fldEmpresa = ""
        fldEstado = ""
        
        On Error GoTo 0
    End If
    
    On Error Resume Next
    imgBrasao.SizeMode = ddSMZoom
    imgLogotipo.SizeMode = ddSMZoom
    On Error GoTo 0
    
    strSql = ""
    strSql = strSql & "SELECT EM.intLogotipo, EM.intBrasao, "
    strSql = strSql & "EM.strNome, UF.strEstado "
    strSql = strSql & "FROM "
    strSql = strSql & gstrEmpresa & " EM, "
    strSql = strSql & gstrUF & " UF "
'    strSql = strSql & "WHERE UF.PKId =* EM.intUF"
    strSql = strSql & "WHERE UF.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & "EM.intUF"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoEmpresa) Then
        With adoEmpresa
            If .EOF = False Then
                LeImagem Val(gstrENulo(!intLogotipo)), imgLogotipo
                LeImagem Val(gstrENulo(!intBrasao)), imgBrasao
                If fldEmpresa Is Nothing = False Then
                    
                    On Error Resume Next
                    
                    fldEmpresa = gstrStringCripitografada(gstrENulo(!STRNOME))
                    'fldEstado = gstrENulo(!strEstado)
                    fldEstado = ""
                    
                    On Error GoTo 0
                    
                End If
            End If
        End With
        adoEmpresa.Close
        Set adoEmpresa = Nothing
        Set gobjBanco = Nothing
    Else
        Exit Sub
    End If
End Sub

Public Sub LeImagem(intCodigo As Integer, _
                    imgImagem, _
           Optional imgImagemGrande As Image)
    '---------------------------------------------------------------
    ' FUNÇÃO USADA PARA MOSTRAR FOTO DO FUNCIONÁRIO.
    '---------------------------------------------------------------
    ' PARÂMETROS:
    ' 1 - intCodigo(Código - Tipo Integer)
    ' 2 - imgImagem(Imagem - Tipo Image)
    ' 3 - imgImagemGrande(Image - Tipo Image)
    '---------------------------------------------------------------

    Dim adoResultado    As ADODB.Recordset
    Dim strSql          As String
    Dim intNumArquivo   As Integer
    Dim intFragmento    As Variant
    Dim bytPedaco()     As Byte
    On Error GoTo ErroLeImagem
    Screen.MousePointer = vbHourglass
    strSql = "SELECT * FROM " & gstrImagem & " WHERE PKId = " & intCodigo
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If .EOF Then
                imgImagem.Picture = LoadPicture()
            Else
                intNumArquivo = FreeFile
                Open App.Path & "\PicTemp" For Binary Access Write As intNumArquivo
                intFragmento = !imgImagem.ActualSize
                bytPedaco() = !imgImagem.GetChunk(intFragmento)
                Put intNumArquivo, , bytPedaco()
                imgImagem.Picture = LoadPicture(App.Path & "\PicTemp")
                Close intNumArquivo
            End If
        End With
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErroLeImagem:
    Resume FimLeImagem:

FimLeImagem:
    Screen.MousePointer = vbDefault
End Sub

'Public Sub VerificaMudancaDeLinhaNoTFBGrid(frmForm As Form, _
'                                           tbb_Grid As TrueOleDBGrid70, _
'                                           gstrTabela As String, _
'                                           blnAlterando As Boolean, _
'                                           blnClickOk As Boolean, _
'                                           objAplicar As Object)
'    With tbb_Grid
'        If (Not .EOF And Not .BOF) And mblnClickOk Then
'            mblnClickOk = False
'            frmForm.txtPKId = .Columns("PKID").Value
'            LeDaTabelaParaObj gstrTabela, frmForm
'            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
'            If objAplicar Is Nothing Then
'                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
'            Else
'                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
'            End If
'            mblnAlterando = True
'            .MarqueeStyle = dbgHighlightRow
'        End If
'    End With
'End Sub
'

Public Sub PeriodoNoRelatorio(frmRelatorio As Form, _
                              lbl_Periodo As Label, _
                              objRelatorio As ActiveReport)
    With frmRelatorio
        lbl_Periodo = "Período: " & .txtdtmInicial & " à " & .txtdtmFinal
        objRelatorio.Caption = .Caption
    End With
End Sub

Public Sub AbreOpcoesExportacao(mobjRelatorio As ActiveReport)
    ExportaRelParaArquivo mobjRelatorio
End Sub

Public Sub MostraEmissorRelatorio(objRelatorio As Object)
Dim objField As Object
On Error Resume Next
    With objRelatorio
        If gbytRelatorioComEmissor = True Then
            .lblEmitido.Visible = True
            .txtEmitido.Visible = True
            .txtEmitido.Text = gstrNomeUsuario
        End If
    End With
     
End Sub

Public Function gstrCidadeConcatenada(vntCidade As Variant, _
                             Optional vntUF As Variant, _
                             Optional vntCEP As Variant) As String
    Dim strAux   As String
    If IsNull(vntCidade) = False Then
        strAux = Trim(vntCidade)
    End If
    If IsMissing(vntUF) = False Then
        If Trim(vntUF) <> "" Then
            If Trim(strAux) <> "" Then
                strAux = strAux & " - " & Trim(vntUF)
            End If
        End If
    End If
    If IsMissing(vntCEP) = False Then
        If Trim(vntCEP) <> "" Then
            If Trim(strAux) <> "" Then
                strAux = gstrCEPFormatado(vntCEP) & " - " & strAux
            End If
        End If
    End If
    gstrCidadeConcatenada = strAux
End Function

Public Function gstrEnderecoConcatenado(vntLogradouro As Variant, _
                               Optional vntTipo As Variant, _
                               Optional vntNumero As Variant, _
                               Optional vntComplemento As Variant, _
                               Optional vntBairro As Variant, _
                               Optional vntTitulo As Variant) As String
    Dim strEndAux   As String
    If IsNull(vntLogradouro) = False Then
        If IsMissing(vntTipo) Then
            If IsMissing(vntTitulo) Then
                strEndAux = Trim(vntLogradouro)
            Else
                strEndAux = Trim(vntTitulo) & " " & Trim(vntLogradouro)
            End If
        ElseIf Trim(vntTipo) = "" Then
            If IsMissing(vntTitulo) Then
                strEndAux = Trim(vntLogradouro)
            Else
                strEndAux = Trim(vntTitulo) & " " & Trim(vntLogradouro)
            End If
        Else
            If IsMissing(vntTitulo) Then
                strEndAux = Trim(vntTipo) & " " & Trim(vntLogradouro)
            Else
                strEndAux = Trim(vntTipo) & " " & Trim(vntTitulo) & " " & Trim(vntLogradouro)
            End If
        End If
        If IsMissing(vntNumero) = False Then
            vntNumero = IIf((IsNull(vntNumero)), 0, vntNumero)
            If Trim(strEndAux) <> "" Then
                If Val(vntNumero) > 0 Then
                    strEndAux = strEndAux & ", " & Val(vntNumero)
                End If
            End If
        End If
        If IsMissing(vntComplemento) = False Then
            If Trim(strEndAux) <> "" Then
                If Trim(vntComplemento) <> "" Then
                    strEndAux = strEndAux & " - " & Trim(vntComplemento)
                End If
            End If
        End If
        If IsMissing(vntBairro) = False Then
            If Trim(strEndAux) <> "" Then
                If Trim(vntBairro) <> "" Then
                    strEndAux = strEndAux & " - " & Trim(vntBairro)
                End If
            End If
        End If
    End If
    gstrEnderecoConcatenado = strEndAux
End Function

Public Function gstrNomeArquivoParaAbrir(Optional strNomeArquivo As String, _
                                         Optional blnSobreEscrevePrompt As Boolean = True, _
                                         Optional strFiltro As String = "All (*.*)| *.*", _
                                         Optional lngIndiceFiltro As Long = 1, _
                                         Optional strDiretorioIncial As String, _
                                         Optional strDlgTitulo As String, _
                                         Optional strExtensaoDefault As String, _
                                         Optional lngProprietario As Long = -1, _
                                         Optional lngFlags As Long) As String
            
    Dim ofnArquivo  As OPENFILENAME
    Dim strAux      As String
    Dim strCaracter As String
    Dim intInd      As Integer
    With ofnArquivo
        .lStructSize = Len(ofnArquivo)
        .flags = (-blnSobreEscrevePrompt * OFN_OVERWRITEPROMPT) Or _
                 OFN_HIDEREADONLY Or _
                 (flags And CLng(Not (OFN_ENABLEHOOK Or _
                                      OFN_ENABLETEMPLATE)))
        If lngProprietario <> -1 Then
            .hWndOwner = lngProprietario
        End If
        .lpstrInitialDir = strDiretorioIncial
        .lpstrDefExt = strExtensaoDefault
        .lpstrTitle = strDlgTitulo
        ' Make new filter with bars (|) replacing nulls and double null at end
        For intInd = 1 To Len(strFiltro)
            strCaracter = Mid$(strFiltro, intInd, 1)
            If strCaracter = "|" Or strCaracter = ":" Then
                strAux = strAux & vbNullChar
            Else
                strAux = strAux & strCaracter
            End If
        Next
        strAux = strAux & vbNullChar & vbNullChar
        .lpstrFilter = strAux
        .nFilterIndex = lngIndiceFiltro
        .lpstrFile = String$(cMaxPath, 0)
        .nMaxFile = cMaxPath
        .lpstrFileTitle = strNomeArquivo & String$(cMaxFile - Len(strNomeArquivo), 0)
        .nMaxFileTitle = cMaxFile
        If GetOpenFileName(ofnArquivo) Then
            gstrNomeArquivoParaAbrir = Mid(.lpstrFile, 1, InStr(.lpstrFile, ".") + 3)
            strNomeArquivo = Mid(.lpstrFileTitle, 1, InStr(.lpstrFileTitle, ".") + 3)
            lngFlags = .flags
            lngIndiceFiltro = .nFilterIndex
            strFiltro = FilterLookup(.lpstrFilter, lngIndiceFiltro)
        Else
            gstrNomeArquivoParaAbrir = sEmpty
            strNomeArquivo = sEmpty
            lngFlags = 0
            lngIndiceFiltro = 0
            strFiltro = sEmpty
        End If
    End With
End Function

Function VBGetOpenFileName(Optional strNomeArquivo As String, _
                           Optional OverWritePrompt As Boolean = True, _
                           Optional Filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional flags As Long) As String
            
    Dim ofnArquivo  As OPENFILENAME
    Dim strAux      As String
    Dim ch          As String
    Dim intInd      As Integer
    With ofnArquivo
        .lStructSize = Len(ofnArquivo)
        .flags = (-OverWritePrompt * OFN_OVERWRITEPROMPT) Or _
                 OFN_HIDEREADONLY Or _
                 (flags And CLng(Not (OFN_ENABLEHOOK Or _
                                      OFN_ENABLETEMPLATE)))
        If Owner <> -1 Then
            .hWndOwner = Owner
        End If
        ' InitDir can take initial directory string
        .lpstrInitialDir = InitDir
        ' DefaultExt can take default extension
        .lpstrDefExt = DefaultExt
        ' DlgTitle can take dialog box title
        .lpstrTitle = DlgTitle
        ' Make new filter with bars (|) replacing nulls and double null at end
        For intInd = 1 To Len(Filter)
            ch = Mid$(Filter, intInd, 1)
            If ch = "|" Or ch = ":" Then
                strAux = strAux & vbNullChar
            Else
                strAux = strAux & ch
            End If
        Next
        strAux = strAux & vbNullChar & vbNullChar
        .lpstrFilter = strAux
        .nFilterIndex = FilterIndex
        .lpstrFile = String$(cMaxPath, 0)
        .nMaxFile = cMaxPath
        .lpstrFileTitle = strNomeArquivo & String$(cMaxFile - Len(strNomeArquivo), 0)
        .nMaxFileTitle = cMaxFile
        If GetOpenFileName(ofnArquivo) Then
            VBGetOpenFileName = Left$(.lpstrFile, Len(.lpstrFile))
            strNomeArquivo = Trim(.lpstrFileTitle)
            flags = .flags
            FilterIndex = .nFilterIndex
            Filter = FilterLookup(.lpstrFilter, FilterIndex)
        Else
            VBGetOpenFileName = sEmpty
            strNomeArquivo = sEmpty
            flags = 0
            FilterIndex = 0
            Filter = sEmpty
        End If
    End With
End Function

Public Sub ProcuraTextoDigitado(intCaracter As Integer, _
                                cboComboBox As ComboBox, _
                       Optional intQtdCampoSaltar As Integer)
    Dim intInd  As Integer
    Dim intAux  As Integer
    If intCaracter = vbKeyReturn Then
        If Trim(cboComboBox.Text) <> "" Then
            For intInd = 0 To cboComboBox.ListCount - 1
                If Trim(cboComboBox.Text) = Trim(cboComboBox.list(intInd)) Then
                    SendKeys "{DOWN}"
                    For intAux = 0 To intQtdCampoSaltar
                        EnviaTeclaTab vbKeyReturn
                    Next
                    Exit Sub
                End If
            Next
        Else
            CaracterValido intCaracter
        End If
    End If
End Sub

Public Function gstrQueryTipoLogradouro(Optional blnRetornaDescricao As Boolean) As String
    
    '--------------------------------------------------------------------------------
    ' FUNCAO USADA PARA MONTAR UMA QUERY ESPECÍFICA PARA
    ' SELECIONAR O TIPO DE LOGRADOURO
    '--------------------------------------------------------------------------------
    ' PARÂMETROS:
    ' 1 - blnParaListView(flag indicando se o objeto a ser preenchido é um ListView)
    '--------------------------------------------------------------------------------

    Dim strSql  As String
    If blnRetornaDescricao Then
        strSql = strSql & "SELECT PKId, strDescricao, strSigla "
        strSql = strSql & "FROM " & gstrTipoLogradouro & " ORDER BY strDescricao"
    Else
        strSql = strSql & "SELECT PKId, strSigla "
        strSql = strSql & "FROM " & gstrTipoLogradouro & " ORDER BY strSigla"
    End If
    gstrQueryTipoLogradouro = strSql
End Function

Public Sub ExportaRelParaArquivo(objRelatorio As ActiveReport)
    Dim strFiltro       As String
    Dim pdf             As New ActiveReportsPDFExport.ARExportPDF
    Dim strNomeArquivo  As String
    On Error GoTo Err_Handle
    strFiltro = ""
    strFiltro = strFiltro & "Rich Text Format (*.RTF)| *.RTF|"
    'strFiltro = strFiltro & "Somente texto (*.TXT)| *.TXT|"
    strFiltro = strFiltro & "Hiper Texto (*.HTML)| *.HTML|"
    strFiltro = strFiltro & "Portable Document Format (*.PDF)| *.PDF|"
    strFiltro = strFiltro & "Planilha do Excel (*.XLS)| *.XLS|"
    If VBGetSaveFileName(strNomeArquivo, "", True, _
        strFiltro, , , "Exportar", "", MDIMenu.hWnd, cdlOFNExplorer _
        Or cdlOFNHideReadOnly Or cdlOFNLongNames) Then
'        If UCase(Mid(strNomeArquivo, InStr(strNomeArquivo, ".") + 1, 3)) = "TXT" Then
'            ExportaParaTXT objRelatorio, strNomeArquivo
'        ElseIf UCase(Mid(strNomeArquivo, InStr(strNomeArquivo, ".") + 1, 3)) = "PDF" Then
'            ExportaParaPDF objRelatorio, strNomeArquivo
'        ElseIf UCase(Mid(strNomeArquivo, InStr(strNomeArquivo, ".") + 1, 3)) = "XLS" Then
'            ExportaParaXLS objRelatorio, strNomeArquivo
'        ElseIf UCase(Mid(strNomeArquivo, InStr(strNomeArquivo, ".") + 1, 3)) = "RTF" Then
'            ExportaParaRTF objRelatorio, strNomeArquivo
'        ElseIf UCase(Mid(strNomeArquivo, InStr(strNomeArquivo, ".") + 1, 4)) = "HTML" Then
'            ExportaParaHTML objRelatorio, strNomeArquivo
'        End If
        If UCase(Right(strFiltro, 3)) = "TXT" Then
            ExportaParaTXT objRelatorio, strNomeArquivo
        ElseIf UCase(Right(strFiltro, 3)) = "PDF" Then
            ExportaParaPDF objRelatorio, strNomeArquivo
        ElseIf UCase(Right(strFiltro, 3)) = "XLS" Then
            ExportaParaXLS objRelatorio, strNomeArquivo
        ElseIf UCase(Right(strFiltro, 3)) = "RTF" Then
            ExportaParaRTF objRelatorio, strNomeArquivo
        ElseIf UCase(Right(strFiltro, 4)) = "HTML" Then
            ExportaParaHTML objRelatorio, strNomeArquivo
        End If
    End If
    Exit Sub
Err_Handle:
    ExibeDetalheErro ""
End Sub

Function VBGetSaveFileName(Filename As String, _
                           Optional FileTitle As String, _
                           Optional OverWritePrompt As Boolean = True, _
                           Optional Filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional flags As Long) As Boolean
            
    Dim opfile As OPENFILENAME, s As String
With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    .flags = (-OverWritePrompt * OFN_OVERWRITEPROMPT) Or _
             OFN_HIDEREADONLY Or _
             (flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hWndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    ' Make new filter with bars (|) replacing nulls and double null at end
    Dim ch As String, i As Integer
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    s = Filename & String$(cMaxPath - Len(Filename), 0)
    .lpstrFile = s
    .nMaxFile = cMaxPath
    s = FileTitle & String$(cMaxFile - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = cMaxFile
    ' All other fields zero
    
    If GetSaveFileName(opfile) Then
        VBGetSaveFileName = True
        Filename = Left$(.lpstrFile, Len(.lpstrFile))
        FileTitle = Left$(.lpstrFileTitle, Len(.lpstrFileTitle))
        flags = .flags
        ' Return the filter index
        FilterIndex = .nFilterIndex
        ' Look up the filter the user selected and return that
        Filter = FilterLookup(.lpstrFilter, FilterIndex)
    Else
        VBGetSaveFileName = False
        Filename = sEmpty
        FileTitle = sEmpty
        flags = 0
        FilterIndex = 0
        Filter = sEmpty
    End If
End With
End Function



Public Sub ExportaParaTXT(rptRelatorio As ActiveReport, strArquivo As String)
    Dim arteArquivo As New ActiveReportsTextExport.ARExportText
    On Error GoTo ErroExportaParaTXT
    arteArquivo.PageDelimiter = gstrDelimitadorPagina
    arteArquivo.TextDelimiter = gstrtxtDelimitadorCampo
    arteArquivo.Unicode = True 'gblnUnicode
    arteArquivo.SuppressEmptyLines = gblnSuprimeLinhaEmBranco
    arteArquivo.Filename = strArquivo
    If rptRelatorio.Pages.Count > 0 Then
        arteArquivo.Export rptRelatorio.Pages
    End If
    Set arteArquivo = Nothing
    Exit Sub
ErroExportaParaTXT:
    Resume FimExportaParaTXT
    
FimExportaParaTXT:

End Sub


Public Sub ExportaParaPDF(rptRelatorio As ActiveReport, strArquivo As String)
    Dim artpArquivo As New ActiveReportsPDFExport.ARExportPDF
    On Error GoTo ErroExportaParaPDF
    artpArquivo.Filename = strArquivo
    If rptRelatorio.Pages.Count > 0 Then
        artpArquivo.Export rptRelatorio.Pages
    End If
    Set artpArquivo = Nothing
    Exit Sub
ErroExportaParaPDF:
    Resume FimExportaParaPDF
    
FimExportaParaPDF:

End Sub

Public Sub ExportaParaRTF(rptRelatorio As ActiveReport, strArquivo As String)
    Dim artrArquivo As New ActiveReportsRTFExport.ARExportRTF
    On Error GoTo ErroExportaParaRTF
    artrArquivo.Filename = strArquivo
    If rptRelatorio.Pages.Count > 0 Then
        artrArquivo.Export rptRelatorio.Pages
    End If
    Set artrArquivo = Nothing
    Exit Sub
    
ErroExportaParaRTF:
    Resume FimExportaParaRTF
    
FimExportaParaRTF:

End Sub

Public Sub ExportaParaHTML(rptRelatorio As ActiveReport, strArquivo As String)
    Dim artxArquivo As New ActiveReportsHTMLExport.HTMLexport
    On Error GoTo ErroExportaParaHTML
    artxArquivo.Filename = strArquivo
    If rptRelatorio.Pages.Count > 0 Then
        artxArquivo.Export rptRelatorio.Pages
    End If
    Set artxArquivo = Nothing
    Exit Sub
    
ErroExportaParaHTML:
    Resume FimExportaParaHTML
    
FimExportaParaHTML:

End Sub

Public Sub ExportaParaXLS(rptRelatorio As ActiveReport, strArquivo As String)
    Dim artxArquivo     As New ActiveReportsExcelExport.ARExportExcel
    On Error GoTo ErroExportaParaXLS
    
    artxArquivo.Filename = strArquivo
    artxArquivo.GenPagebreaks = True
    artxArquivo.TrimEmptySpace = False
    artxArquivo.ShowMarginSpace = True
    If rptRelatorio.Pages.Count > 0 Then
        artxArquivo.Export rptRelatorio.Pages
    End If
    
    Set artxArquivo = Nothing
    Exit Sub
    
ErroExportaParaXLS:
    Resume FimExportaParaXLS
    
FimExportaParaXLS:

End Sub

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
    Dim iStart As Long, iEnd As Long, s As String
    iStart = 1
    If sFilters = sEmpty Then Exit Function
    Do
        ' Cut out both parts marked by null character
        iEnd = InStr(iStart, sFilters, vbNullChar)
        If iEnd = 0 Then Exit Function
        iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
        If iEnd Then
            s = Mid$(sFilters, iStart, iEnd - iStart)
        Else
            s = Mid$(sFilters, iStart)
        End If
        iStart = iEnd + 1
        If iCur = 1 Then
            FilterLookup = s
            Exit Function
        End If
        iCur = iCur - 1
    Loop While iCur
End Function

Sub Configura_Relatorio(rptRelatorio As ActiveReport, Optional blnSetupDialog As Boolean, Optional blnConfigMargem As Boolean = True)
    Dim strSql              As String
    Dim blnFlag             As Boolean
    
    
    On Error GoTo Err_Handle
  
    If Not blnConfigMargem Then
    
        rptRelatorio.Printer.PaperSize = gintTamanhoDoPapel
    
        blnFlag = True

        If blnSetupDialog Then
            blnFlag = rptRelatorio.Printer.SetupDialog
            gintTamanhoDoPapel = rptRelatorio.Printer.PaperSize
            strSql = ""
            strSql = "Update " & gstrUsuarios & " Set intTamanhoDoPapel = " & gintTamanhoDoPapel & " "
            strSql = strSql & "Where PKId = " & glngCodUsr
            Set gobjBanco = New clsBanco
            gobjBanco.Execute strSql
        End If
    Else
        blnFlag = rptRelatorio.PageSetup(MDIMenu.hWnd)
        intOrientacao = rptRelatorio.Printer.Orientation
        gblnRestartRelatorio = True
        rptRelatorio.Restart
        
        blnFlag = False
    End If
    
    If blnFlag Then
        If rptRelatorio.Version = "2.0" Then
            rptRelatorio.PageSettings.LeftMargin = 900
            rptRelatorio.PageSettings.RightMargin = 500
            If (rptRelatorio.PrintWidth + rptRelatorio.PageSettings.LeftMargin + rptRelatorio.PageSettings.RightMargin) > rptRelatorio.Printer.PaperWidth Then
                Do While (rptRelatorio.PrintWidth + rptRelatorio.PageSettings.LeftMargin + rptRelatorio.PageSettings.RightMargin) > rptRelatorio.Printer.PaperWidth
                    If rptRelatorio.PageSettings.RightMargin > 2 Then
                        rptRelatorio.PageSettings.RightMargin = rptRelatorio.PageSettings.RightMargin - 2
                    ElseIf rptRelatorio.PageSettings.LeftMargin > 2 Then
                        rptRelatorio.PageSettings.LeftMargin = rptRelatorio.PageSettings.LeftMargin - 2
                    Else
                        Exit Do
                    End If
                Loop
            End If
        Else
            rptRelatorio.PageLeftMargin = 900
            rptRelatorio.PageRightMargin = 500
            If (rptRelatorio.PrintWidth + rptRelatorio.PageLeftMargin + rptRelatorio.PageRightMargin) > rptRelatorio.Printer.PaperWidth Then
                Do While (rptRelatorio.PrintWidth + rptRelatorio.PageLeftMargin + rptRelatorio.PageRightMargin) > rptRelatorio.Printer.PaperWidth
                    If rptRelatorio.PageRightMargin > 2 Then
                        rptRelatorio.PageRightMargin = rptRelatorio.PageRightMargin - 2
                    ElseIf rptRelatorio.PageLeftMargin > 2 Then
                        rptRelatorio.PageLeftMargin = rptRelatorio.PageLeftMargin - 2
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        
        If blnSetupDialog Then
            gblnRestartRelatorio = True
            rptRelatorio.Restart
        End If
    End If
Exit Sub
Err_Handle:
    ExibeDetalheErro ""
End Sub

Sub TrocaCorParaZebrado(vntDetalhe As Variant, Optional objDetalhe As Object)
    
    '---------------------------------------------------------------------------
    ' SUB USADA PARA TROCAR A COR DO OBJETO DE FUNDO DOS RELATÓRIOS
    ' TORNADO O RELATÓRIO COM LISTA (ZEBRADO)
    '---------------------------------------------------------------------------
    ' PARÂMETROS:
    ' 1 - vntDetalhe (o objeto (label) do fundo do relatório)
    '---------------------------------------------------------------------------

    Static blnFlag          As Boolean
    
    If gblnRelatorioZebrado Then
        With vntDetalhe
            If Not blnFlag Then
                .BackColor = gvntCorZebrado
                blnFlag = True
            Else
                .BackColor = vbWhite
                blnFlag = False
            End If
        End With
    Else
        vntDetalhe.BackColor = vbWhite
    End If
    
    
    On Error Resume Next
    If Not IsMissing(objDetalhe) Then
        vntDetalhe.Height = objDetalhe.Height
    End If
    
End Sub

Sub TrocaCorDaSecaoParaZebrado(vntSecao As Variant)
    
    '---------------------------------------------------------------------------
    ' SUB USADA PARA TROCAR A COR DO OBJETO DE FUNDO DOS RELATÓRIOS
    ' TORNADO O RELATÓRIO COM LISTA (ZEBRADO)
    '---------------------------------------------------------------------------
    ' PARÂMETROS:
    ' 1 - vntDetalhe (o objeto (label) do fundo do relatório)
    '---------------------------------------------------------------------------

    Static blnFlag    As Boolean
    If gblnRelatorioZebrado Then
        With vntSecao
            .BackStyle = 1
            If Not blnFlag Then
                .BackColor = gvntCorZebrado
                blnFlag = True
            Else
                .BackColor = vbWhite
                blnFlag = False
            End If
        End With
    Else
        vntSecao.BackColor = vbWhite
    End If
End Sub

Public Function gstrParteInteira(vntValor1 As Variant, vntValor2 As Variant) As Variant
    Dim vntVarAux   As Variant
    Dim vntValor3   As Variant
    Dim lngInd      As Long
    vntValor3 = CDec(gstrConvVrDoSql(vntValor1 / vntValor2, 15))
    For lngInd = 1 To Len(vntValor3)
        If Mid(vntValor3, lngInd, 1) = "," Then
            Exit For
        End If
        vntVarAux = vntVarAux & Mid(vntValor3, lngInd, 1)
    Next
    gstrParteInteira = vntVarAux
End Function

Public Function gvntPeriodoReal(vntDataOuNumero As Variant, _
                       Optional vntDataFinal As Variant, _
                       Optional strIntervalo As String, _
                       Optional intAnos As Integer, _
                       Optional bytMeses As Byte, _
                       Optional bytDias As Byte)
    '---------------------------------------------------------------
    ' FUNÇÃO USADA PARA CALCULAR O NÚMERO DE ANOS, MESES E DIAS
    ' A PARTIR DE UM NÚMERO OU ENTRE DUAS DATAS
    '---------------------------------------------------------------
    ' PARÂMETROS
    ' Esta função recebe e/ou retorna os seguintes parâmetros, nesta ordem
    ' A) parâmetros recebidos:
    ' 1 - Data Inicial ou número
    ' 2 - Data final
    ' 3 - Intervalo ('D" dias, 'M' mese ou 'A' anos)
    ' Obs: 1 - o primeiro parâmetro pode ser uma data ou um número
    '          natural e tem que ser informado, caso não
    '          seja informado a função retona um campo nulo
    '      2 - o segundo parâmetro pode ser omitido. Se esse for
    '          omitido e o primeriro parâmetro for uma data a função
    '          usa a data do servidor para calcular o número de dias
    '      3 - o terceiro parâmetro pode ser omitido, e, se for omitido
    '          a função retorna uma string contendo: o número de dias,
    '          meses e anos (ex: 5 anos, 10 meses e 22 dias).
    ' B) parâmetros retornados:
    ' 4 - Número de anos no perído
    ' 5 - Número de meses no perído
    ' 5 - Número de dias restantes no perído
    ' Obs: 2 - os parâmetros de 4 a 6 podem ser omitidos
    '---------------------------------------------------------------
    Const strDia As String = " Dia"
    Const strDias As String = " Dias"
    Const strAno As String = " Ano"
    Const strAnos As String = " Anos"
    Const strMes As String = " Mês"
    Const strMeses As String = " Meses"
    Dim dblQuociente    As Variant
    Dim dblResto        As Variant
    Dim bytInd          As Byte
    Dim strPeriodo      As String
    'Verifica tipo e validade dos parâmetros
    If IsDate(vntDataOuNumero) Then
        If IsDate(vntDataFinal) Then
            vntDataOuNumero = DateDiff("d", vntDataOuNumero, vntDataFinal)
        Else
            vntDataOuNumero = DateDiff("d", vntDataOuNumero, gstrDataDoSistema)
        End If
    End If
    If IsNumeric(vntDataOuNumero) = False Or Val(vntDataOuNumero) < 1 Then
        Exit Function
    End If
    'Verifica o intervalo e calcula os anos, meses e dias
    If UCase(Trim(strIntervalo)) = "D" Then
        gvntPeriodoReal = vntDataOuNumero
        Exit Function
    ElseIf UCase(Trim(strIntervalo)) = "M" Then
        gvntPeriodoReal = vntDataOuNumero \ 30.4375
        Exit Function
    Else
        dblResto = vntDataOuNumero / 365.25 - gstrParteInteira(vntDataOuNumero, 365.25)
        intAnos = CDec(gstrConvVrDoSql(vntDataOuNumero / 365.25, 15)) - dblResto
        dblQuociente = dblResto * 365.25
        dblResto = CDec(gstrConvVrDoSql(dblQuociente / 30.4375, 15)) - gstrParteInteira(dblQuociente, 30.4375)
        bytMeses = dblQuociente / 30.4375 - dblResto
        If Format(dblResto * 30.4375, "00") > 30 Or Format(dblResto * 30.4375, "00") = 30 Then
            bytMeses = bytMeses + 1
        ElseIf dblResto > 0 Then
            bytDias = dblResto * 30.4375
        Else
            bytDias = 0
        End If
        If bytMeses = 12 Then
            intAnos = intAnos + 1
            bytMeses = 0
        End If
        If UCase(Trim(strIntervalo)) = "A" Then
            gvntPeriodoReal = intAnos
            Exit Function
        End If
    End If
    'Monta uma string com os anos, meses e dias encontrados
    '-------------------------------------------------------
    If intAnos > 0 Then
        If intAnos > 1 Then
            strPeriodo = intAnos & strAnos
        Else
            strPeriodo = intAnos & strAno
        End If
    End If
    If bytMeses > 0 Then
        If strPeriodo <> "" Then
            If bytDias = 0 Then
                If bytMeses > 1 Then
                    strPeriodo = strPeriodo & " e " & bytMeses & strMeses
                Else
                    strPeriodo = strPeriodo & " e " & bytMeses & strMes
                End If
            Else
                If bytMeses > 1 Then
                    strPeriodo = strPeriodo & ", " & bytMeses & strMeses
                Else
                    strPeriodo = strPeriodo & ", " & bytMeses & strMes
                End If
            End If
        Else
            If bytMeses > 1 Then
                strPeriodo = bytMeses & strMeses
            Else
                strPeriodo = bytMeses & strMes
            End If
        End If
    End If
    If bytDias > 0 Then
        If strPeriodo <> "" Then
            If bytDias > 1 Then
                strPeriodo = strPeriodo & " e " & bytDias & strDias
            Else
                strPeriodo = strPeriodo & " e " & bytDias & strDia
            End If
        Else
            If bytDias > 1 Then
                strPeriodo = strPeriodo & bytDias & strDias
            Else
                strPeriodo = strPeriodo & bytDias & strDia
            End If
        End If
    End If
    gvntPeriodoReal = strPeriodo
End Function

Public Function gbytZeroOuUm(vntFlag As Variant) As Byte
    If UCase(vntFlag) = "SIM" Or UCase(vntFlag) = "M" Then
        gbytZeroOuUm = 1
    ElseIf UCase(vntFlag) = "NÃO" Or UCase(vntFlag) = "F" Then
        gbytZeroOuUm = 0
    Else
        gbytZeroOuUm = Val(vntFlag)
    End If
End Function

Sub MoveItemNoListView(objLista As ListView, flgParaCima As Boolean)
    Dim intSelecionado  As Integer
    Dim intInd          As Integer
    Dim strAux          As String
    Dim lngTagAux       As Long
    With objLista
        If flgParaCima Then
            If .SelectedItem.Index < .ListItems.Count Then
                intSelecionado = .SelectedItem.Index
                strAux = .ListItems(intSelecionado + 1).Text
                .ListItems(intSelecionado + 1).Text = .ListItems(intSelecionado).Text
                .ListItems(intSelecionado).Text = strAux
                For intInd = 1 To .ColumnHeaders.Count - 1
                    strAux = .ListItems(intSelecionado + 1).SubItems(intInd)
                    .ListItems(intSelecionado + 1).SubItems(intInd) = .ListItems(intSelecionado).SubItems(intInd)
                    .ListItems(intSelecionado).SubItems(intInd) = strAux
                Next
                lngTagAux = .ListItems(intSelecionado + 1).Tag
                .ListItems(intSelecionado + 1).Tag = .ListItems(intSelecionado).Tag
                .ListItems(intSelecionado).Tag = lngTagAux
                .ListItems(intSelecionado).Selected = False
                .ListItems(intSelecionado + 1).Selected = True
            End If
        Else
            If .SelectedItem.Index > 1 Then
                intSelecionado = .SelectedItem.Index
                strAux = .ListItems(intSelecionado - 1).Text
                .ListItems(intSelecionado - 1).Text = .ListItems(intSelecionado).Text
                .ListItems(intSelecionado).Text = strAux
                For intInd = 1 To .ColumnHeaders.Count - 1
                    strAux = .ListItems(intSelecionado - 1).SubItems(intInd)
                    .ListItems(intSelecionado - 1).SubItems(intInd) = .ListItems(intSelecionado).SubItems(intInd)
                    .ListItems(intSelecionado).SubItems(intInd) = strAux
                Next
                lngTagAux = .ListItems(intSelecionado - 1).Tag
                .ListItems(intSelecionado - 1).Tag = .ListItems(intSelecionado).Tag
                .ListItems(intSelecionado).Tag = lngTagAux
                .ListItems(intSelecionado).Selected = False
                .ListItems(intSelecionado - 1).Selected = True
            End If
        End If
    End With
End Sub


Public Function gstrTotalDeRegistros(strTabela As String, _
                                    lblLabel As Object) As String
                                    
    '------------------------------------------------------------------------------'
    '      FUNCTION USADA PARA RETORNAR O NUMERO DE REGISTROS DE UM RECORDSET      '
    '------------------------------------------------------------------------------'
    ' PARÂMETRO:                                                                   '
    ' strTabela (Tabela de onde será lido o registro)                              '
    ' lblLabel (Label que tras o nome do tipo de registro a ser lido no relatório.)'
    '------------------------------------------------------------------------------'

    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    Dim intQT           As Integer

    strSql = ""
    strSql = strSql & "SELECT COUNT(*) as SOMA FROM " & strTabela
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            intQT = adoResultado!SOMA
        End If
    End If
    If intQT <> 0 Then
        gstrTotalDeRegistros = "Totalização de " & lblLabel.Caption & " : " & intQT & " registro(s)."
        Else
        gstrTotalDeRegistros = ""
        MsgBox "Nenhum registro foi encontrado."
    End If
End Function

Public Function gstrPalavraComPrimeiraMaiuscula(ByVal strTexto As String) As String
    Dim vntPalavra      As Variant
    Dim intIndTabela    As Integer
    vntPalavra = Split(strTexto, " ")
    strTexto = ""
    For intIndTabela = 0 To UBound(vntPalavra)
        Select Case Trim(UCase(vntPalavra(intIndTabela)))
        Case "A", "E", "O", "À", "AS", "OS", "AO", "OU", "DA", "DE", "DO", "DAS", "DOS"
            If intIndTabela = 0 Then
                strTexto = strTexto & Trim(UCase(Mid(vntPalavra(intIndTabela), 1, 1)))
                strTexto = strTexto & Trim(LCase(Mid(vntPalavra(intIndTabela), 2)))
            ElseIf Trim(strTexto) <> "" Then
                strTexto = strTexto & " " & Trim(LCase(vntPalavra(intIndTabela)))
            Else
                strTexto = Trim(LCase(vntPalavra(intIndTabela)))
            End If
        Case Else
            If Trim(strTexto) <> "" And Trim(vntPalavra(intIndTabela)) <> "" Then
                strTexto = strTexto & " " & Trim(UCase(Mid(vntPalavra(intIndTabela), 1, 1)))
            Else
                strTexto = strTexto & Trim(UCase(Mid(vntPalavra(intIndTabela), 1, 1)))
            End If
            strTexto = strTexto & Trim(LCase(Mid(vntPalavra(intIndTabela), 2)))
        End Select
    Next
    gstrPalavraComPrimeiraMaiuscula = strTexto
End Function

Public Sub Limpa_Controles(frmForm As Form, bLimpaTXT As Boolean, _
                           blLimpaCHK As Boolean, blLimpaOPT As Boolean, _
                           blLimpaDBC As Boolean, blLimpaLVW As Boolean)
    '------------------------------------------'
    ' ESTA SUB LIMPA O CONTEÚDO DOS CONTROLES  '
    ' DO FORMULÁRIO.                           '
    '------------------------------------------'
    ' PARÂMETROS:                              '
    '                                          '
    ' 1 - Formulário                           '
    ' 2 - Indica se limpa TextBox(Boolean)     '
    ' 3 - Indica se limpa CheckBox(Boolean)    '
    ' 4 - Indica se limpa OptionButton(Boolean)'
    ' 5 - Indica se limpa DataCombo(Boolean)    '
    ' 6 - Indica se limpa ListView(Boolean)    '
    '------------------------------------------'
    Dim iCountCtr As Integer
    For iCountCtr = 0 To frmForm.Controls.Count - 1
        If TypeOf frmForm.Controls(iCountCtr) Is TextBox And bLimpaTXT Then
            frmForm.Controls(iCountCtr).Text = ""
        ElseIf TypeOf frmForm.Controls(iCountCtr) Is CheckBox And blLimpaCHK Then
            frmForm.Controls(iCountCtr).Value = 0
        ElseIf TypeOf frmForm.Controls(iCountCtr) Is OptionButton And blLimpaOPT Then
            frmForm.Controls(iCountCtr).Value = False
        ElseIf TypeOf frmForm.Controls(iCountCtr) Is DataCombo And blLimpaDBC Then
            frmForm.Controls(iCountCtr).BoundText = ""
        ElseIf TypeOf frmForm.Controls(iCountCtr) Is ListView And blLimpaLVW Then
            frmForm.Controls(iCountCtr).ListItems.Clear
        End If
    Next
End Sub

Public Sub gCorLinhaSelecionada(tdb_Grid As TDBGrid)
With tdb_Grid
'Ultima do gusm
.MarqueeStyle = dbgHighlightRow
.HighlightRowStyle.BackColor = vbHighlight
.HighlightRowStyle.ForeColor = vbWhite

End With
End Sub

'==========='==========='==========='==========='==========='==========='==========='==========='==========='===========
'==========='==========='==========='==========='==========='==========='==========='==========='==========='===========
'==========='==========='==========='===== BARRA DE FERRAMENTAS '==== '============='==========='==========='===========
'==========='==========='==========='=====       DEFAULT        '==== '============='==========='==========='===========
'==========='==========='==========='==========='==========='==========='==========='==========='==========='===========
'==========='==========='==========='==========='==========='==========='==========='==========='==========='===========

Public Sub UpdateToolbar(ab As ActiveBar2)
Dim i As Integer
Dim j As Integer

Dim tool As ActiveBar2LibraryCtl.tool
Dim iCat As Integer
Dim keys(0) As New ShortCut

ab.DisplayKeysInToolTip = True
ab.DisplayToolTips = True

CreateTools ab
CreateBands ab
CreateTools1 ab
CreateBands1 ab

ab.PersonalizedMenus = ddPMDisabled

AjustaPersonalizar ab

'On Error Resume Next
'If ab.Version = "2.5.0.22" Then
'    ab.XPLook = True
'    ab.PersonalizedMenus = ddPMDisplayOnHover
'End If
'On Error GoTo 0

ab.RecalcLayout
ab.Refresh

End Sub

Private Sub CreateTools(ab As ActiveBar2)
    Dim tool    As ActiveBar2LibraryCtl.tool
    Dim iCat    As Integer
    Dim keys(0) As New ShortCut

    iCat = 4000
    Set tool = ab.Tools.Add(iCat + 1, gstrMnuArquivo)
    tool.Caption = "&Arquivo"
    tool.SubBand = gstrMnuArquivo
    tool.Category = gstrMnuArquivo
    
    Set tool = ab.Tools.Add(iCat + 2, "mnuEdit")
    tool.Caption = "&Editar"
    tool.SubBand = "mnuEdit"
    tool.Category = gstrMnuArquivo
    
    Set tool = ab.Tools.Add(iCat + 3, "mnuView")
    tool.Caption = "&Exibir"
    tool.SubBand = "mnuView"
    tool.Category = gstrMnuArquivo
    
    Set tool = ab.Tools.Add(iCat + 4, "mnuHelp")
    tool.Caption = "&Ajuda"
    tool.SubBand = "mnuHelp"
    tool.Category = gstrMnuArquivo
    
    Set tool = ab.Tools.Add(iCat + 20, "mnuWindow")
    tool.Caption = "&Janela"
    tool.SubBand = "mnuWindow"
    tool.Category = gstrMnuArquivo
    
'================ Arquivo
    'iCat = 2000
    Set tool = ab.Tools.Add(iCat + 5, gstrNovo)
    tool.Caption = "&Novo": tool.Category = gstrMnuArquivo
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    keys(0) = "Control+N"
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrNovo).Picture
    tool.ShortCuts = keys
    tool.ToolTipText = "Novo"
    
    Set tool = ab.Tools.Add(iCat + 6, gstrSalvar)
    tool.Caption = "&Salvar": tool.Category = gstrMnuArquivo
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrSalvar).Picture
    keys(0) = "Control+S"
    tool.ShortCuts = keys
    tool.ToolTipText = "Salvar"
    
    Set tool = ab.Tools.Add(iCat + 7, gstrImprimir)
    tool.Caption = "&Imprimir": tool.Category = gstrMnuArquivo
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrImprimir).Picture
    keys(0) = "Control+P"
    tool.ShortCuts = keys
    tool.ToolTipText = "Imprimir"
    
    Set tool = ab.Tools.Add(iCat + 8, gstrDeletar)
    tool.Caption = "&Excluir": tool.Category = gstrMnuArquivo
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrDeletar).Picture
    keys(0) = "Alt+F5"
    tool.ShortCuts = keys
    tool.ToolTipText = "Excluir"

    
    Set tool = ab.Tools.Add(iCat + 9, gstrAplicar)
    tool.Caption = "&Aplicar": tool.Category = gstrMnuArquivo
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrAplicar).Picture
    tool.Enabled = False
    keys(0) = "Alt+F6"
    tool.ShortCuts = keys
    tool.ToolTipText = "Aplicar"
    
    Set tool = ab.Tools.Add(iCat + 10, gstrLocalizar)
    tool.Caption = "&Localizar": tool.Category = gstrMnuArquivo
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrLocalizar).Picture
    tool.Enabled = True
    keys(0) = "F5"
    tool.ShortCuts = keys
    tool.ToolTipText = "Localizar"
    
    Set tool = ab.Tools.Add(iCat + 11, gstrPreencherLista)
    tool.Caption = "Preencher &lista": tool.Category = gstrMnuArquivo
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrPreencherLista).Picture
    tool.Enabled = True
    keys(0) = "F6"
    tool.ShortCuts = keys
    tool.ToolTipText = "Preencher lista de opções"
    
    
    Set tool = ab.Tools.Add(iCat + 12, gstrFechar)
    tool.Caption = "Fechar": tool.Category = gstrMnuArquivo
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrFechar).Picture
    keys(0) = "Control+F4"
    tool.ShortCuts = keys
    tool.ToolTipText = "Fechar"

    Set tool = ab.Tools.Add(iCat + 13, "SAIR")
    tool.Caption = "Sai&r": tool.Category = gstrMnuArquivo
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    keys(0) = "Alt+F4"
    tool.ShortCuts = keys
    tool.ToolTipText = "Sair"

'=========== Editar
    'iCat = 3000
    Set tool = ab.Tools.Add(iCat + 14, "miECut")
    tool.Caption = "Recor&tar": tool.Category = "Edit"
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages("RECORTAR").Picture
    keys(0).Clear
    tool.ShortCuts = keys
    
    Set tool = ab.Tools.Add(iCat + 15, "miECopy")
    tool.Caption = "&Copiar": tool.Category = "Edit"
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages("COPIAR").Picture
    'keys(0) = "Control+C"
    tool.ShortCuts = keys
    
    Set tool = ab.Tools.Add(iCat + 16, "miEPaste")
    tool.Caption = "C&olar": tool.Category = "Edit"
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages("COLAR").Picture
    'keys(0) = "Control+V"
    tool.ShortCuts = keys
    
'=========== Visualizar
    'iCat = 4000
    Set tool = ab.Tools.Add(iCat + 17, "miVToolbar")
    tool.Caption = "&Barra de Ferramentas": tool.Category = "View"
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    tool.Checked = True
    
    Set tool = ab.Tools.Add(iCat + 18, "miVStatusBar")
    tool.Caption = "&Status": tool.Category = "View"
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    tool.Checked = True
    
'========== Ajuda
    'iCat = 5000
    Set tool = ab.Tools.Add(iCat + 19, "miHContents")
    tool.Caption = "&Conteúdo": tool.Category = "Help"
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages("AJUDA").Picture
    keys(0) = "F1"
    tool.ShortCuts = keys

'========== Window
    'iCat = 6000
    Set tool = ab.Tools.Add(iCat + 21, "miHorizontal")
    tool.Caption = "Organizar &Horizontalmente": tool.Category = "Janelas"
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    
    Set tool = ab.Tools.Add(iCat + 22, "miVertical")
    tool.Caption = "Organizar &Verticalmente": tool.Category = "Janelas"
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    
    Set tool = ab.Tools.Add(iCat + 23, "miCascata")
    tool.Caption = "&Cascata": tool.Category = "Janelas"
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed

    Set tool = ab.Tools.Add(iCat + 24, "miJanelas")
    tool.Caption = "&Janelas": tool.Category = "Janelas"
    tool.MenuVisibility = ddMVVisibleIfRecentlyUsed
    tool.ControlType = ddTTWindowList

    'Set Tool = ab.Tools.Add(iCat + 20, "miHTecnical")
    'Tool.Caption = "&Suporte Técnico": Tool.Category = "Help"
    'Tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages("SUPORTE").Picture
    
End Sub


Private Sub CreateBands(ab As ActiveBar2)
    Dim b As ActiveBar2LibraryCtl.Band
    Dim intID       As Integer
    Set b = ab.Bands.Add(gstrMnuArquivo): b.Type = ddBTPopup
    b.flags = 95
    With b.Tools
        .Insert .Count, ab.Tools(gstrNovo)
        .Insert .Count, ab.Tools(gstrSalvar)
        .Insert .Count, ab.Tools(gstrDeletar)
        .Insert .Count, ab.Tools(gstrAplicar)
        .Insert .Count, ab.Tools(gstrImprimir)
        .Insert .Count, ab.Tools(gstrLocalizar)
        .Insert .Count, ab.Tools(gstrPreencherLista)
        .Insert .Count, ab.Tools(gstrFechar)
        .Insert .Count, ab.Tools("SAIR")
    End With


    Set b = ab.Bands.Add("mnuEdit"): b.Type = ddBTPopup
    ab.Bands("mnuEdit").Visible = False
    b.flags = 95
    With b.Tools
        .Insert .Count, ab.Tools("miECut")
        .Insert .Count, ab.Tools("miECopy")
        .Insert .Count, ab.Tools("miEPaste")
    End With


    Set b = ab.Bands.Add("mnuView"): b.Type = ddBTPopup
    b.flags = 95
    With b.Tools
    .Insert .Count, ab.Tools("miVToolbar")
    .Insert .Count, ab.Tools("miVStatusBar")
    End With


    Set b = ab.Bands.Add("mnuHelp"): b.Type = ddBTPopup
    b.flags = 95
    b.Tools.Insert b.Tools.Count, ab.Tools("miHContents")
    
 
    Set b = ab.Bands.Add("mnuWindow"): b.Type = ddBTPopup
    b.flags = 95
    b.Tools.Insert b.Tools.Count, ab.Tools("miHorizontal")
    b.Tools.Insert b.Tools.Count, ab.Tools("miVertical")
    b.Tools.Insert b.Tools.Count, ab.Tools("miCascata")
    b.Tools.Insert b.Tools.Count, ab.Tools("miJanelas")
    
    'B.Tools.Insert B.Tools.Count, ab.Tools("miHTecnical")
    
 
    Set b = ab.Bands.Add("mnuMain"): b.Type = ddBTMenuBar
    b.flags = 95
    b.Caption = ""
    ab.Tools(gstrMnuArquivo).SubBand = gstrMnuArquivo
    ab.Tools("mnuEdit").SubBand = "mnuEdit"
    ab.Tools("mnuView").SubBand = "mnuView"
    ab.Tools("mnuHelp").SubBand = "mnuHelp"
    ab.Tools("mnuWindow").SubBand = "mnuWindow"

    With b.Tools
        .Insert .Count, ab.Tools(gstrMnuArquivo)
        .Insert .Count, ab.Tools("mnuEdit")
        .Insert .Count, ab.Tools("mnuView")
        .Insert .Count, ab.Tools("mnuHelp")
        .Insert .Count, ab.Tools("mnuWindow")

    End With
    
    ab.RecalcLayout
    ab.Refresh
End Sub


Private Sub CreateTools1(ab As ActiveBar2)
    Dim tool    As ActiveBar2LibraryCtl.tool
    Dim iCat    As Integer
    Dim keys(0) As New ShortCut
    
    iCat = 4000
    Set tool = ab.Tools.Add(iCat + 1, gstrBtnArquivo)
    tool.Caption = "&Arquivo": tool.SubBand = gstrBtnArquivo: tool.Category = gstrBtnArquivo
    
    Set tool = ab.Tools.Add(iCat + 2, "1mnuEdit")
    tool.Caption = "&Editar": tool.SubBand = "mnuEdit": tool.Category = "1Menus"
    
    Set tool = ab.Tools.Add(iCat + 3, "1mnuView")
    tool.Caption = "&Exibir": tool.SubBand = "mnuView": tool.Category = "1Menus"
    
    Set tool = ab.Tools.Add(iCat + 4, "1mnuHelp")
    tool.Caption = "&Ajuda": tool.SubBand = "mnuHelp": tool.Category = "1Menus"

'========= mnuFile
    'iCat = 2000
    Set tool = ab.Tools.Add(iCat + 5, gstrNovo)
    tool.Caption = "&Novo": tool.Category = gstrBtnArquivo
    keys(0) = "Control+N"
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrNovo).Picture
    tool.ShortCuts = keys
    tool.ToolTipText = "Novo"
    
    Set tool = ab.Tools.Add(iCat + 6, gstrSalvar)
    tool.Caption = "&Salvar": tool.Category = gstrBtnArquivo
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrSalvar).Picture
    keys(0) = "Control+S"
    tool.ShortCuts = keys
    tool.ToolTipText = "Salvar"
    
    Set tool = ab.Tools.Add(iCat + 7, gstrImprimir)
    tool.Caption = "&Imprimir": tool.Category = gstrBtnArquivo
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrImprimir).Picture
    keys(0) = "Control+P"
    tool.ShortCuts = keys
    tool.ToolTipText = "Imprimir"
    
    Set tool = ab.Tools.Add(iCat + 8, gstrDeletar)
    tool.Caption = "&Excluir": tool.Category = gstrBtnArquivo
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrDeletar).Picture
    keys(0) = "Alt+F5"
    tool.ShortCuts = keys
    tool.ToolTipText = "Excluir"

    Set tool = ab.Tools.Add(iCat + 9, gstrAplicar)
    tool.Caption = "&Aplicar": tool.Category = gstrMnuArquivo
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrAplicar).Picture
    keys(0) = "Alt+F6"
    tool.ShortCuts = keys
    
    Set tool = ab.Tools.Add(iCat + 10, gstrLocalizar)
    tool.Caption = "&Localizar": tool.Category = gstrMnuArquivo
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrLocalizar).Picture
    keys(0) = "F5"
    tool.ShortCuts = keys
    
    Set tool = ab.Tools.Add(iCat + 11, gstrPreencherLista)
    tool.Caption = "&Preencher &lista": tool.Category = gstrMnuArquivo
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrPreencherLista).Picture
    keys(0) = "F6"
    tool.ShortCuts = keys
    
    Set tool = ab.Tools.Add(iCat + 12, gstrFechar)
    tool.Caption = "&Fechar": tool.Category = gstrBtnArquivo
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages(gstrFechar).Picture
    keys(0) = "Control+F4"
    tool.ShortCuts = keys
    tool.ToolTipText = "Fechar"

'========== mnuEdit
    'iCat = 3000
    Set tool = ab.Tools.Add(iCat + 14, "1miECut")
    tool.Caption = "Recor&tar": tool.Category = "1Edit"
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages("RECORTAR").Picture
    keys(0).Clear
    tool.ShortCuts = keys
    tool.ToolTipText = "Recortar"
    
    Set tool = ab.Tools.Add(iCat + 15, "1miECopy")
    tool.Caption = "&Copiar": tool.Category = "1Edit"
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages("COPIAR").Picture
    'keys(0) = ""
    tool.ShortCuts = keys
    tool.ToolTipText = "Copiar"
    
    Set tool = ab.Tools.Add(iCat + 16, "1miEPaste")
    tool.Caption = "&C&olar": tool.Category = "1Edit"
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages("COLAR").Picture
    'keys(0) = ""
    tool.ShortCuts = keys
    tool.ToolTipText = "Colar"
    
'========== mnuView
    'iCat = 4000
    Set tool = ab.Tools.Add(iCat + 17, "1miVToolbar")
    tool.Caption = "&Barra de Ferramentas": tool.Category = "1View"
    tool.Checked = True
    
    Set tool = ab.Tools.Add(iCat + 18, "1miVStatusBar")
    tool.Caption = "&Status": tool.Category = "1View"
    tool.Checked = True
    
'========== mnuHelp
    'iCat = 5000
    Set tool = ab.Tools.Add(iCat + 19, "1miHContents")
    tool.Caption = "&Conteúdo": tool.Category = "1Help"
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages("AJUDA").Picture
    keys(0) = "F1"
    tool.ShortCuts = keys
    tool.ToolTipText = "Ajuda"

    'Set Tool = ab.Tools.Add(iCat + 20, "1miHTecnical")
    'Tool.Caption = "&Suporte Técnico": Tool.Category = "1Help"
    'Tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages("SUPORTE").Picture
    'Tool.ToolTipText = "Suporte Técnico"
End Sub

Private Sub CreateBands1(ab As ActiveBar2)
    Dim acbBandeira As ActiveBar2LibraryCtl.Band
    Set acbBandeira = ab.Bands.Add(gstrBtnArquivo)
    
    acbBandeira.Type = ddBTNormal
    acbBandeira.Caption = "Arquivo"
    acbBandeira.flags = 95
    acbBandeira.DisplayMoreToolsButton = False
    
    With acbBandeira.Tools
        .Insert .Count, ab.Tools(gstrNovo)
        .Insert .Count, ab.Tools(gstrSalvar)
        .Insert .Count, ab.Tools(gstrDeletar)
        .Insert .Count, ab.Tools(gstrAplicar)
        .Insert .Count, ab.Tools(gstrImprimir)
        .Insert .Count, ab.Tools(gstrLocalizar)
        .Insert .Count, ab.Tools(gstrPreencherLista)
        .Insert .Count, ab.Tools(gstrFechar)
    End With
    
    Set acbBandeira = ab.Bands.Add("1mnuEdit")
    acbBandeira.Type = ddBTNormal
    acbBandeira.Caption = "Editar"
    acbBandeira.flags = 95
    acbBandeira.DisplayMoreToolsButton = False
    
    With acbBandeira.Tools
        .Insert .Count, ab.Tools("1miECut")
        .Insert .Count, ab.Tools("1miECopy")
        .Insert .Count, ab.Tools("1miEPaste")
    End With
    
    Set acbBandeira = ab.Bands.Add("1mnuHelp")
    acbBandeira.Type = ddBTNormal
    acbBandeira.Caption = "Ajuda"
    acbBandeira.flags = 95
    acbBandeira.DisplayMoreToolsButton = False
    
    With acbBandeira.Tools
        .Insert acbBandeira.Tools.Count, ab.Tools("1miHContents")
        '.Insert acbBandeira.Tools.Count, ab.Tools("1miHTecnical")
    End With
    ab.RecalcLayout
    ab.Refresh
End Sub
'==========='==========='==========='==========='==========='==========='==========='==========='==========='===========
'==========='==========='==========='==========='==========='==========='==========='==========='==========='===========
'==========='==========='==========='===== BARRA DE FERRAMENTAS '==== '============='==========='==========='===========
'==========='==========='==========='=====       DEFAULT        '==== '============='==========='==========='===========
'==========='==========='==========='==========='==========='==========='==========='==========='==========='===========
'==========='==========='==========='==========='==========='==========='==========='==========='==========='===========

Public Sub HabilitaDesabilitaBotao1(blnFlag As Boolean, _
                                    strBandeira As String, _
                         ParamArray vntListaBotao() As Variant)
    Dim vntBotao    As Variant
    
    '-------------------------------------------------------------------
    ' SUB USADA PARA HABILTAR OU DESABILITAR BOTOES BARRA DE FERRAMENTA
    '-------------------------------------------------------------------
    ' PARÂMETROS
    ' 1 - blnFlag(verdadeiro ou falso - habilitar ou desabilitar)
    ' 2 - strBandeira(Grupo de botões que serão alterados)
    ' 3 - vntListaBotao() - Botões que serão alterados
    '   OBS: Se o primeiro parâmetro for boleano, todos os botões serão
    '   habilitados/desabilitados mantendo o botão fechar habilitado.
    '   Este parâmetro será usado principalmente
    '   por relatórios que tem sua própria barra de ferramentas.
    '-------------------------------------------------------------------
    
    On Error Resume Next
    
    With MDIMenu.actBarra.Bands(strBandeira)
        .Tools(gstrNovo).Enabled = True
        .Tools(gstrSalvar).Enabled = True
        .Tools(gstrImprimir).Enabled = True
        
        For Each vntBotao In vntListaBotao
            'Primeiro parâmetro
            If vntBotao = True Then
                .Tools(gstrNovo).Enabled = blnFlag
                .Tools(gstrSalvar).Enabled = blnFlag
                .Tools(gstrAplicar).Enabled = blnFlag
                .Tools(gstrDeletar).Enabled = blnFlag
                .Tools(gstrImprimir).Enabled = blnFlag
            Else
                .Tools(vntBotao).Enabled = blnFlag
            End If
        Next
    End With
    gblnVerificaPermissoes gintCodSeguranca, gstrMnuArquivo
    
''    With MDIMenu.actBarra.Bands(strBandeira)
''        If .Tools(gstrNovo).Enabled = False Then
''            For Each vntBotao In vntListaBotao
''                If vntBotao = gstrDeletar Then
''                    If blnFlag = True Then
''                        .Tools(gstrSalvar).Enabled = True
''                    End If
''                End If
''            Next
''        End If
''    End With
    
End Sub

Public Function gblnTrocaNomeBancoDeDados() As Boolean
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "SELECT * "
    strSql = strSql & "FROM " & gstrDatabases
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not (.BOF And .EOF) Then
                gstrFuncionarioRH = Trim(!RH) & ".dbo.Funcionario"
                gstrSecretariaRH = Trim(!RH) & ".dbo.Secretaria"
                gstrDepartamentoRH = Trim(!RH) & ".dbo.Depto"
                gstrSecaoRH = Trim(!RH) & ".dbo.Secao"
                gstrSetorRH = Trim(!RH) & ".dbo.Setor"
                gstrOrgaoRH = Trim(!RH) & ".dbo.Orgao"
                gstrSituacaoRH = Trim(!RH) & ".dbo.Situacao"
                gstrProventoDescontoRH = Trim(!RH) & ".dbo.ProventoDesconto"
                gstrCargoRH = Trim(!RH) & ".dbo.Cargo"
                gstrPensionistaRH = Trim(!RH) & ".dbo.Pensionista"
                gstrEvolucaoValorRH = Trim(!RH) & ".dbo.EvolucaoValor"
                gstrPagamentosRH = Trim(!RH) & ".dbo.Pagamentos"
                gstrBairroRH = Trim(!RH) & ".dbo.Bairro"
                gstrFotoRH = Trim(!RH) & ".dbo.Foto"
                
                gstrBaseADGover = Trim(!Conjunto)
                gstrBaseRH = Trim(!RH)
            End If
        End With
    End If

    gblnTrocaNomeBancoDeDados = True
End Function

Public Sub gGravaHistoricoContribuinte(strCodigo As String, lngContribuinte As Long, _
                              strTransacao As String, strNomeSistema As String, dblValor As Double)
                              '  Grava na tabela tblHistoricoContribuinte os dados abaixo

Dim strSql As String
Dim ado As ADODB.Recordset


strSql = ""
strSql = strSql & "INSERT INTO " & gstrHistoricoContribuinte
strSql = strSql & " (strCodigo, intContribuinte, strTransacao, "
strSql = strSql & "strNomeSistema, dblValor, lngCodUsr)"

strSql = strSql & " VALUES("

strSql = strSql & "'" & strCodigo & "'"
strSql = strSql & "," & lngContribuinte
strSql = strSql & ",'" & strTransacao & "'"
strSql = strSql & ",'" & strNomeSistema & "'"
strSql = strSql & "," & gstrConvVrParaSql(dblValor)
'strsql = strsql & ", GETDATE()"
strSql = strSql & "," & glngCodUsr & ")"
'strsql = strsql & "  ,   GETDATE())"

Set gobjBanco = New clsBanco
gobjBanco.Execute strSql

End Sub

Public Sub CarregaPermissoes()
    
'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    Dim strSistema   As String
    Dim i As Long
    
    On Error GoTo err_CarregaPermissoes
    
    Select Case UCase(App.ProductName)
        Case "TRIBUTARIO"
            strSistema = "J"
        Case "ORCAMENTARIO"
            strSistema = "F"
        Case "FROTA"
            strSistema = "B"
        Case "RH"
            strSistema = "I"
        Case "LEGISLACAO"
            strSistema = "C"
        Case "OUVIDORIA"
            strSistema = "M"
        Case "COMPRAS"
            strSistema = "A"
        Case "PATRIMONIO"
            strSistema = "G"
        Case "MATERIAL"
            strSistema = "E"
        Case "PROTOCOLO"
            strSistema = "H"
        Case "SEGURANCA"
            strSistema = "L"
        Case "MENOR"
            strSistema = "D"
        Case "GERENCIAL"
            strSistema = "N"
    End Select
    
    strSql = ""
    strSql = strSql & "SELECT P.strPermissao, I.intCodigo "
    strSql = strSql & "FROM " & gstrPermissoes & " P, " & gstrItens & " I "
    strSql = strSql & "WHERE P.intItem = I.PKId "
    strSql = strSql & "AND P.intUsuario = " & glngCodUsr & " "
    strSql = strSql & "AND I.blnPermissao = 1 "
'    strSql = strSql & "AND UPPER(SUBSTRING(I.strCodItem,1,1)) = '" & strSistema & "' "
    strSql = strSql & "AND UPPER(" & strSUBSTRING & "(I.strCodItem,1,1)) = '" & strSistema & "' "
    strSql = strSql & "ORDER BY I.intCodigo"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 0, adoResultado) Then
        With adoResultado
            i = 0
            ReDim vetPermissoes(.RecordCount)
            Do While Not .EOF
                i = i + 1
                If i = 10 Then
                   i = 0
                End If
                frmSplash.lblProgresso.Caption = "Verificando permissões do usuário " & String(i, ".")
                frmSplash.lblProgresso.Refresh
                DoEvents
                vetPermissoes(.AbsolutePosition).intCodigo = !intCodigo
                vetPermissoes(.AbsolutePosition).strPermissao = !strPermissao
                .MoveNext
            Loop
        End With
    End If
    Exit Sub
err_CarregaPermissoes:
End Sub

Public Function gblnExisteTabela(STRNOME As String, _
                        Optional blnExibeMensagem As Boolean) As Boolean

'******************************************************************************************
' Data: 09/03/2003
' Alteração: - Adaptação do bloco Transact-SQL (SQL Server) para o PL/SQL (Oracle).
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim adoResultado    As ADODB.Recordset
    Dim strSql          As String
    strSql = ""
    If (bytDBType = EDatabases.SQLServer) Then
        strSql = strSql & "IF EXISTS (SELECT NAME FROM SYSOBJECTS "
        strSql = strSql & "WHERE NAME = '" & STRNOME & "') "
        strSql = strSql & "SELECT Existe = 1 ELSE SELECT Existe = 0"
    
    ElseIf (bytDBType = EDatabases.Oracle) Then
        strSql = strSql & "SELECT COUNT(*) Existe "
        strSql = strSql & "FROM ALL_OBJECTS "
        strSql = strSql & "WHERE OBJECT_NAME = '" & STRNOME & "' "
        strSql = strSql & "AND UPPER(OWNER) = 'CPDMASTER'"
    
    End If
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If !existe = 1 Then
                gblnExisteTabela = True
            ElseIf blnExibeMensagem Then
                ExibeMensagem "A tabela " & STRNOME & " não foi encontrada"
            End If
        .Close
        End With
    End If
End Function

Public Function VerificaPermissoesBotao(objObjeto As Object)
    Dim objControle As Object
    
    On Error Resume Next
    
    For Each objControle In objObjeto.Controls
        If LCase(Mid(objControle.Name, 1, 4)) = "cmd_" Then
            If Val(objControle.Tag) <> 0 Then
                If gstrPermissao(objControle.Tag) <> "" Then
                    If InStr(1, gstrPermissao(objControle.Tag), "2") = 0 Then
                        objControle.Enabled = False
                    End If
                End If
            End If
        End If
    Next
End Function

Public Function gblnValidaSenha(intCaracter As Integer) As Boolean
    Select Case intCaracter
        Case vbKeyBack, vbKeyDelete
        Case 48 To 59, 65 To 90, 97 To 122
        Case Else
            gblnValidaSenha = False
            Exit Function
    End Select
    gblnValidaSenha = True
End Function

Public Function gblnSistemaDemonstracao(strTabela As String, lngQuantidade As Long) As Boolean
    Dim adoResultado As ADODB.Recordset
    Dim strSql       As String
    
    On Error GoTo err_gblnSistemaDemonstracao
        
    gblnSistemaDemonstracao = False

    If Not gblnDemonstracao Then
        Exit Function
    End If
    
    strSql = ""
    strSql = "SELECT COUNT(*) AS Total FROM " & strTabela
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not (.BOF And Not .EOF) Then
                If !TOTAL >= lngQuantidade Then
                    frmRegistro.Show 1
                    gblnSistemaDemonstracao = True
                    Exit Function
                End If
            End If
        End With
    End If
err_gblnSistemaDemonstracao:
End Function

Public Function gstrQueryRelatorioGuiaDeArrecadacao(ByRef blnExiteLancamento As Boolean, strInscricaoInicial As String, _
                                                    strInscricaoFinal As String, intExercicio As Integer, intComposicaoDaReceita As Integer, _
                                                    Optional blnTodasIncricoes As Boolean, Optional dtmDataVencimento As Date, _
                                                    Optional intParcelaInicial As Integer, Optional intParcelaFinal As Integer) As String

'******************************************************************************************
' Data: 06/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 24/03/2003
' Alteração: - Retirado o comando CONVERT da cláusula SELECT uma vez que este não era
'            necessário.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset

'By Power "C.strComplementoC, 'CEP.: ' + CONVERT(NVarChar,intCepC) AS intCepC, strBairroC, " & _

    strSql = "SELECT I.PKId,  I.intExercicio, J.intNumeroParcela, I.strInscricaoCadastral, " & _
            "R.intCodigo CodReceita, J.dtmDataVencimento , C.strNome AS Contribuinte, "
'            "E.strDescricao + ' - ' + F.strSigla AS Municipio , C.strLogradouroC AS Logradouro, C.intNumeroC, " & _
'            "C.strComplementoC, 'CEP.: ' + CONVERT(VarChar,intCepC) AS intCepC, strBairroC, "
    strSql = strSql & _
            "E.strDescricao" & strCONCAT & "' - '" & strCONCAT & "F.strSigla AS Municipio , C.strLogradouroC AS Logradouro, C.intNumeroC, " & _
            "C.strComplementoC, 'CEP.: '" & strCONCAT & gstrCONVERT(CDT_VARCHAR, "intCepC") & " AS intCepC, strBairroC, " & _
            "J.PKId PKIdParcelaReceita, J.dblJuros, J.dblMulta " & _
            "FROM " & _
            gstrParcelaReceita & " J, " & _
            gstrContribuinte & " C, " & _
            gstrLancamentoCalculo & " I, " & _
            gstrComposicaoDaReceita & " R, " & _
            gstrCidade & " E, " & _
            gstrUF & " F " & _
            "WHERE " & _
            "e.PKId = c.intMunicipioC " & _
            "AND R.PKId = J.intComposicaoDaReceita " & _
            "AND J.intLancamentoCalculo = I.PKId " & _
            "AND I.intContribuinte = C.PKId " & _
            "AND F.PKId = C.intUFC"
    If Not blnTodasIncricoes Then
        strSql = strSql & " AND I.strInscricaoCadastral BETWEEN '" & strInscricaoInicial & "' AND '" & strInscricaoFinal & "' "
        
    End If
    strSql = strSql & " AND I.intExercicio = " & intExercicio & _
            " AND J.intComposicaoDaReceita = " & intComposicaoDaReceita
    If dtmDataVencimento <> 0 Then
        strSql = strSql & " AND J.dtmDataVencimento = " & gstrConvDtParaSql(dtmDataVencimento)
    End If
'    If intParcelaInicial <> 0 And intParcelaFinal <> 0 Then
        strSql = strSql & " AND J.intNumeroParcela BETWEEN " & intParcelaInicial & " AND " & intParcelaFinal
'    End If

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            blnExiteLancamento = True
        Else
            ExibeMensagem "Não Existe(m) Lançamento(s) para a(s) Inscrição Selecionada(s)"
            blnExiteLancamento = False
        End If
    End If

    gstrQueryRelatorioGuiaDeArrecadacao = strSql
End Function

Public Function bytRetornaCodigoModulo(strModulo As String)
Dim bytCodigo As Byte

Select Case UCase(Trim(strModulo))
Case "COMPRAS"
    bytCodigo = 1
Case "CONCURSO PUBLICO"
    bytCodigo = 2
Case "MENOR"
    bytCodigo = 3
Case "ESCOLAR"
    bytCodigo = 4
Case "FROTA"
    bytCodigo = 5
Case "LEGISLACAO"
    bytCodigo = 6
Case "MATERIAL"
    bytCodigo = 7
Case "ORCAMENTARIO"
    bytCodigo = 8
Case "OUVIDORIA"
    bytCodigo = 9
Case "PATRIMONIO"
    bytCodigo = 10
Case "PROTOCOLO"
    bytCodigo = 11
Case "RH"
    bytCodigo = 12
Case "SEGURANCA"
    bytCodigo = 13
Case "TRIBUTARIO"
    bytCodigo = 14
End Select
bytRetornaCodigoModulo = bytCodigo
End Function

Sub AlwaysOnTop(FormName As Form, bOnTop As Boolean)
'Coloca o form "always on top"
    Dim Success As Integer
    If bOnTop = False Then
        Success% = SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    Else
        Success% = SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
    End If
End Sub

Public Sub ExibeDetalheErro(Optional strMsg As String, _
                            Optional strQuery As String, _
                            Optional Icone As IconesErro)
    Load frmDetalheErro
    If Trim(strMsg) <> "" Then
        frmDetalheErro.lbl_Mensagem1 = Trim(strMsg)
    End If
    If Trim(strQuery) <> "" Then
        frmDetalheErro.txt_Query = strQuery
    End If
    If Icone <> 0 Then
        With frmDetalheErro
            If Icone <= .img_Imagens.ListImages.Count Then
                .img_Erro.Picture = .img_Imagens.ListImages(Icone).Picture
            End If
        End With
    End If
   If frmDetalheErro.Visible = False Then
       frmDetalheErro.Show vbModal
    End If
    
End Sub

Public Function glngDiasBackup() As Long
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    
    On Error GoTo err_glngDiasBackup
    
    strSql = ""
    strSql = strSql & "SELECT * "
    strSql = strSql & "FROM " & gstrBackup
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not (.BOF And .EOF) Then
                glngDiasBackup = DateDiff("d", CDate(gstrDataFormatada(!dtmBackup)), CDate(gstrDataDoSistema))
            Else
                glngDiasBackup = -1
            End If
        End With
    Else
        glngDiasBackup = -1
    End If
err_glngDiasBackup:
End Function

Public Sub PreencherListaDeOpcoes(objObjeto As Object, Optional vntPKId As Variant)
    Dim strSql              As String
    Dim strAlternativa      As String
    Dim i                   As Integer
    Dim iAlternativo        As Integer
    Dim strCondicao         As String
    Dim strValor            As String
    Dim strCampo            As String
    Dim strCampo2           As String
    Dim strCampoChave       As String
    Dim vetTag()            As String
    Dim strAux              As String
    Dim blnCompras          As Boolean
    Dim adoResultado        As ADODB.Recordset
    On Error GoTo err_PreencherListaDeOpcoes
    blnCompras = False
    If objObjeto Is Nothing Then Exit Sub
        
    If Not TypeOf objObjeto Is DataCombo Then
        Exit Sub
    ElseIf objObjeto.Style <> 0 Then
        Exit Sub
    ElseIf Trim(objObjeto.Tag) = "" Then
        Exit Sub
    End If
    
    vetTag = Split(Trim(objObjeto.Tag), ";")
    
    If InStr(objObjeto.Tag, gstrContribuinte) Then
    
        If UCase(App.ProductName) = "COMPRAS" Then
            blnCompras = True
            strAlternativa = "SELECT " 'DISTINCT "
            strAlternativa = strAlternativa & Right(Left(objObjeto.Tag, InStr(objObjeto.Tag, gstrContribuinte) - 1), Len(Left(objObjeto.Tag, InStr(objObjeto.Tag, gstrContribuinte) - 1)) - 7)
            strAlternativa = strAlternativa & Mid(objObjeto.Tag, InStr(objObjeto.Tag, gstrContribuinte), Len(gstrContribuinte))
            strAlternativa = strAlternativa & " CO, " & gstrItens & " IT, " & gstrModuloContribuinte & " MC "
            If InStr(objObjeto.Tag, "WHERE") Then
                strAlternativa = strAlternativa & Mid(objObjeto.Tag, InStrRev(UCase(objObjeto.Tag), "WHERE"), IIf(InStr(UCase(objObjeto.Tag), "ORDER") > 0, InStrRev(UCase(objObjeto.Tag), "ORDER"), Len(objObjeto.Tag)) - InStrRev(UCase(objObjeto.Tag), "WHERE")) & " AND "
            Else
                strAlternativa = strAlternativa & " WHERE "
            End If
            strAlternativa = strAlternativa & " IT.PKId = MC.intItem AND "
            strAlternativa = strAlternativa & "MC.intContribuinte = CO.PKId AND "
            strAlternativa = strAlternativa & "IT.PkId IN (SELECT PKid FROM " & gstrItens & " WHERE "
            ''If gintModulo = 8 Then
            ''    strAlternativa = strAlternativa & "UPPER(stritem)= 'MATERIAIS')" 'OR
            ''ElseIf gintModulo = 9 Then
            ''    strAlternativa = strAlternativa & "UPPER(stritem)='ORÇAMENTÁRIO')" ' OR
            ''ElseIf gintModulo = 1 Then
                strAlternativa = strAlternativa & "UPPER(stritem)='COMPRAS')"
            ''End If
        Else
            strAlternativa = Left(objObjeto.Tag, InStrRev(objObjeto.Tag, gstrContribuinte) - 1)
            strAlternativa = strAlternativa & Mid(objObjeto.Tag, InStrRev(objObjeto.Tag, gstrContribuinte), Len(gstrContribuinte))
            strAlternativa = strAlternativa & " CO, " & gstrItens & " IT, " & gstrModuloContribuinte & " MC "
            If InStr(objObjeto.Tag, "WHERE") Then
                strAlternativa = strAlternativa & Mid(objObjeto.Tag, InStrRev(UCase(objObjeto.Tag), "WHERE"), IIf(InStr(UCase(objObjeto.Tag), "ORDER") > 0, InStrRev(UCase(objObjeto.Tag), "ORDER"), Len(objObjeto.Tag)) - InStrRev(UCase(objObjeto.Tag), "WHERE")) & " AND "
            Else
                strAlternativa = strAlternativa & " WHERE "
            End If
            strAlternativa = strAlternativa & "IT.PKId = MC.intItem AND "
            strAlternativa = strAlternativa & "MC.intContribuinte = CO.PKId AND "
            strAlternativa = strAlternativa & "IT.PkId = " & gintModulo
        End If
        
        strAlternativa = strAlternativa & " AND CO.BLNINATIVO = 0"
        If InStr(objObjeto.Tag, "ORDER") Then
            strAlternativa = strAlternativa & " " & Mid(objObjeto.Tag, InStrRev(UCase(objObjeto.Tag), "ORDER"), IIf(InStr(objObjeto.Tag, ";"), InStr(objObjeto.Tag, ";") - InStrRev(UCase(objObjeto.Tag), "ORDER"), Len(objObjeto.Tag)))
        Else
            strAlternativa = strAlternativa & " ORDER BY CO.strNome"
        End If
        If blnCompras = True Then
            strAux = Left(strAlternativa, InStrRev(strAlternativa, "(SELECT") - 1)
            
            strAux = Replace(UCase(strAux), " PKID", " CO.PKId")
            strAux = Replace(UCase(strAux), " STRNOME", " CO.strNome")
            strAlternativa = strAux & Right(strAlternativa, Len(strAlternativa) - InStrRev(strAlternativa, "(SELECT") + 1)
        Else
            strAux = Left(strAlternativa, InStrRev(strAlternativa, "FROM") - 1)
            
            strAux = Replace(UCase(strAux), "PKID", "CO.PKId")
            strAux = Replace(UCase(strAux), "STRNOME", "CO.strNome")
            strAlternativa = strAux & Right(strAlternativa, Len(strAlternativa) - InStrRev(strAlternativa, "FROM") + 1)
        End If
        
        
    End If
    
    strSql = Trim(vetTag(0)) 'Query a ser utilizada
    
    strCampo = Trim(vetTag(1)) 'Nome do Campo a ser consultado
    'Se houver mais do que um campo de busca
    If InStr(InStr(1, objObjeto.Tag, ";") + 1, objObjeto.Tag, ";") > 0 Then strCampo2 = Trim(vetTag(2)) 'Nome do outro campo a ser consultado
    
    vetTag = Split(Trim(objObjeto.Tag), " ")
    strCampoChave = Trim(vetTag(1)) 'Nome do Campo chave
    If UCase(strCampoChave) = "DISTINCT" Or InStr(1, UCase(strCampoChave), "MAX(") Then
        strCampoChave = Trim(vetTag(2)) 'Nome do Campo chave
    End If
    If InStr(1, strCampoChave, ",") <> 0 Then
        strCampoChave = Mid(strCampoChave, 1, InStr(1, strCampoChave, ",") - 1)
    ElseIf InStr(1, strCampoChave, ")") <> 0 Then
        strCampoChave = Mid(strCampoChave, 1, InStr(1, strCampoChave, ")") - 1)
    End If
    
    If strSql = "" Or strCampo = "" Then
        Exit Sub
    End If
    
    If Not IsMissing(vntPKId) Then
        strValor = Val(vntPKId)
        strCondicao = strCampoChave & " = " & strValor
    ElseIf objObjeto.MatchedWithList = True Then  'Existe item selecionado
        'Rafael - 23/09/2004
        If Val(objObjeto.BoundText) = gintPkidFixo Then 'Usado para o Select do Tag com Pkid fixo
           strValor = Trim(objObjeto.Text)
           'Verificacao de campo no caso de Inscricao ou Aviso, para completar com zeros a esquerda
           If UCase(strCampo) = "STRINSCRICAO" Or UCase(strCampo) = "STRINSCRICAOCADASTRAL" Then
               strValor = String(gintLenInscricao - Len(Trim(strValor)), "0") & strValor
           End If
           If UCase(strCampo) = "STRNUMEROAVISO" Then
                strValor = String(gintLenNumAviso - Len(Trim(strValor)), "0") & strValor
           End If
            
           If UCase(strCampo) = "STRINSCRICAO" Or UCase(strCampo) = "STRINSCRICAOCADASTRAL" Or UCase(strCampo) = "STRNUMEROAVISO" Then
                strCondicao = " " & strCampo & " LIKE '" & UCase(strValor) & "%'"
           Else
                strCondicao = "UPPER(" & strCampo & ") LIKE '" & UCase(strValor) & "%'"
           End If
           
           If Trim(strCampo2) <> "" Then strCondicao = "(" & strCondicao & " OR " & strCampo2 & " LIKE '" & strValor & "%')"
        Else
            strValor = Val(objObjeto.BoundText)
            If UCase(strCampo) = "STRINSCRICAO" Or UCase(strCampo) = "STRINSCRICAOCADASTRAL" Or UCase(strCampo) = "STRNUMEROAVISO" Then
                strCondicao = " " & strCampoChave & " = " & UCase(strValor)
            Else
                strCondicao = "UPPER(" & strCampoChave & ") = " & UCase(strValor)
            End If
            
        End If
    Else
        If Trim(objObjeto.BoundText) <> "" Then
            strValor = Trim(objObjeto.BoundText)
            'Verificacao de campo no caso de Inscricao ou Aviso, para completar com zeros a esquerda
            If UCase(strCampo) = "STRINSCRICAO" Or UCase(strCampo) = "STRINSCRICAOCADASTRAL" Then
                strValor = String(gintLenInscricao - Len(Trim(strValor)), "0") & strValor
            End If
            If UCase(strCampo) = "STRNUMEROAVISO" Then
                strValor = String(gintLenNumAviso - Len(Trim(strValor)), "0") & strValor
            End If
            
            
            If UCase(strCampo) = "STRINSCRICAO" Or UCase(strCampo) = "STRINSCRICAOCADASTRAL" Or UCase(strCampo) = "STRNUMEROAVISO" Then
               
                 strCondicao = " " & strCampo & " LIKE '" & UCase(strValor) & "%'"
            Else
                strCondicao = "UPPER(" & strCampo & ") LIKE '" & UCase(strValor) & "%'"
            End If
            
            If Trim(strCampo2) <> "" Then strCondicao = "(" & strCondicao & " OR " & strCampo2 & " LIKE '" & strValor & "%')"
        End If
    End If

    If strCondicao <> "" Then
        i = InStr(1, UCase(strSql), "WHERE")
        
        If strAlternativa <> "" Then iAlternativo = InStr(1, UCase(strAlternativa), "WHERE")
        
        If (strAlternativa = "" And i <> 0) Or (strAlternativa <> "" And iAlternativo <> 0) Then
            If strAlternativa <> "" Then
                strAlternativa = Mid(strAlternativa, 1, iAlternativo + 5) & Replace(UCase(Replace(UCase(strCondicao), "STRNOME", "CO.strNome")), "PKID", "CO.PKId") & " AND " & Mid(strAlternativa, iAlternativo + 5)
            End If
            If i > 0 Then
                strSql = Mid(strSql, 1, i + 5) & strCondicao & " AND " & Mid(strSql, i + 5)
            End If
        Else
            i = InStr(1, UCase(strSql), "GROUP BY")
            If i <> 0 Then
                If strAlternativa <> "" Then
                    strAlternativa = Mid(strAlternativa, 1, iAlternativo - 1) & "WHERE " & Replace(UCase(Replace(UCase(strCondicao), "STRNOME", "CO.strNome")), "PKID", "CO.PKId") & " " & Mid(strAlternativa, iAlternativo)
                End If
                strSql = Mid(strSql, 1, i - 1) & "WHERE " & strCondicao & " " & Mid(strSql, i)
            Else
                i = InStr(1, UCase(strSql), "ORDER BY")
                If i <> 0 Then
                    If strAlternativa <> "" Then
                        strAlternativa = Mid(strAlternativa, 1, iAlternativo - 1) & "WHERE " & Replace(UCase(Replace(UCase(strCondicao), "STRNOME", "CO.strNome")), "PKID", "CO.PKId") & " " & Mid(strAlternativa, iAlternativo)
                    End If
                    strSql = Mid(strSql, 1, i - 1) & "WHERE " & strCondicao & " " & Mid(strSql, i)
                Else
                    If strAlternativa <> "" Then
                        strAlternativa = strAlternativa & " WHERE " & strCondicao
                        strAlternativa = Mid(strAlternativa, 1, iAlternativo - 1) & "WHERE " & Replace(UCase(Replace(UCase(strCondicao), "STRNOME", "CO.strNome")), "PKID", "CO.PKId") & " " & Mid(strAlternativa, iAlternativo)
                    End If
                    strSql = strSql & " WHERE " & strCondicao
                End If
            End If
        End If
    End If
    Screen.MousePointer = 11
    
    If strAlternativa <> "" Then
    
        Set adoResultado = New ADODB.Recordset
        
        On Error Resume Next
        adoResultado.Open strAlternativa, gcncADOMain, adOpenForwardOnly, adLockReadOnly
        
        If Err.Number = 0 Or Err.Number = -2147217871 Then
            LeDaTabelaParaObj "", objObjeto, strAlternativa
        Else
            LeDaTabelaParaObj "", objObjeto, strSql
        End If
        
        adoResultado.Close
        
        On Error GoTo 0
        
        Set adoResultado = Nothing
    Else
        LeDaTabelaParaObj "", objObjeto, strSql
    End If
    
    If Not IsMissing(vntPKId) Then
        objObjeto.BoundText = strValor
    End If
    Screen.MousePointer = 0
    Exit Sub
err_PreencherListaDeOpcoes:
    Screen.MousePointer = 0
    ExibeDetalheErro ""
End Sub

Public Function gstrQueryDataComboBairro()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM " & gstrBairro & " "
    strSql = strSql & "ORDER BY strDescricao"
    gstrQueryDataComboBairro = strSql
End Function

Public Function gstrQueryDataComboUF()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strSigla "
    strSql = strSql & "FROM " & gstrUF & " "
    strSql = strSql & "ORDER BY strSigla"
    gstrQueryDataComboUF = strSql
End Function

Public Function gstrQueryDataComboMunicipio()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM " & gstrCidade & " "
    strSql = strSql & "ORDER BY strDescricao"
    gstrQueryDataComboMunicipio = strSql
End Function

Public Function gstrQueryDataComboTipoLogradouro()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strSigla "
    strSql = strSql & "FROM " & gstrTipoLogradouro & " "
    strSql = strSql & "ORDER BY strSigla"
    gstrQueryDataComboTipoLogradouro = strSql
End Function

Public Function gstrQueryDataComboTituloLogradouro()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM " & gstrTituloLogradouro & " "
    strSql = strSql & "ORDER BY strDescricao"
    gstrQueryDataComboTituloLogradouro = strSql
End Function

Public Function gstrQueryDataComboBanco()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM " & gstrBanco & " "
    strSql = strSql & "ORDER BY strDescricao"
    gstrQueryDataComboBanco = strSql
End Function

Public Sub DropDownDataCombo(ByRef DataCombo As DataCombo, ByRef dForm As Form, _
                                  Optional Area As Integer = -1, Optional KeyCode As Integer = -1, _
                                  Optional Shift As Integer = -1)
    
    Dim strText     As String
    
    DataCombo.MatchEntry = dblExtendedMatching
    
    If Area <> -1 Then
        If (Area = 0) And (Trim(DataCombo.Text) <> "") And (Not DataCombo.MatchedWithList) Then
            dForm.MantemForm gstrPreencherLista
        End If
    Else
        If (KeyCode = 40 And Shift = 4) Or (KeyCode = 115 And Shift = 0) Then
            If Trim(DataCombo.Text) <> "" And Not DataCombo.MatchedWithList Then
                dForm.MantemForm gstrPreencherLista
            End If
        ElseIf KeyCode = 13 Then
            dForm.MantemForm gstrPreencherLista
        End If
    End If
    
End Sub

Public Function VerificaVersao() As Boolean
Dim strLocalizacaoExecutavel            As String
Dim strAplicacao                        As String
Dim objFileSystem                       As Scripting.FileSystemObject
Dim objFolder                           As Scripting.Folder
Dim objFiles                            As Scripting.Files
Dim objFile                             As Scripting.file
Dim objFileSystemAtual                  As Scripting.FileSystemObject
Dim objFileAtual                        As Scripting.file

On Error GoTo Problema_VerificaVersao
    
    VerificaVersao = True
    
    strLocalizacaoExecutavel = gstrLeValorRegister(HKEY_CURRENT_USER, _
                                        "SOFTWARE\CPD\AdGover\Parâmetros", "Atualizador")
    
    If strLocalizacaoExecutavel = "" Then VerificaVersao = False
    
    strAplicacao = App.EXEName & ".exe"
        
    Set objFileSystem = New Scripting.FileSystemObject
    
    Set objFileSystemAtual = New Scripting.FileSystemObject
   
    Set objFolder = objFileSystem.GetFolder(strLocalizacaoExecutavel & "\")
    
    Set objFileAtual = objFileSystem.GetFile(App.Path & "\" & strAplicacao)
       
    Set objFiles = objFolder.Files
        
    For Each objFile In objFiles
    
        If UCase(Trim(objFile.Name)) = UCase(Trim(strAplicacao)) Then
            
            If CDate(objFileAtual.DateLastModified) < CDate(objFile.DateLastModified) Then
                ExibeMensagem "Existe uma nova versão disponível. O WinPublic Update será executado."
                Shell App.Path & "\AdGoverUpdate.exe " & strAplicacao, vbNormalFocus
                End
            End If
            
            Exit For
            
        End If
    Next
    
Problema_VerificaVersao:

    If Err.Number <> 0 Then
        ExibeMensagem "Não foi possivel procurar por versões mais atuais do sistema. Por favor, entre em contato com o administrador do sistema" & vbCrLf & vbCrLf & "Descrição do erro: " & Err.Description
    End If
    
End Function

Public Function gstrProximoCodigo(txtDestino As Object, _
                              strTabela As String, _
                              strCampo As String, _
                              intCodigo As Integer, _
                              Optional strGrupo As String, _
                              Optional strValorGrupo As String, _
                              Optional intMascaraEspecifica As Integer, _
                              Optional Retorno As Boolean, _
                              Optional strSubGrupo As String, _
                              Optional strValorSubGrupo As String, _
                              Optional strParametroEspecifico As String, _
                              Optional strValorParametroEspecifico As String, _
                              Optional bitAutoNumeracaoParaMsgCriticas As Byte = 0) As String
                              

'******************************************************************************************
' Data: 22/04/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql          As String
Dim adoResultado    As ADODB.Recordset
Dim intOrdenacao    As Integer
Dim strSistema      As String
Dim intFor          As Integer
   
    If Not Retorno Then If txtDestino <> "" Then Exit Function
    
    Set gobjBanco = New clsBanco
    
    Select Case UCase(App.ProductName)
        Case "TRIBUTARIO"
            strSistema = "J"
        Case "ORCAMENTARIO"
            strSistema = "F"
        Case "FROTA"
            strSistema = "B"
        Case "RH"
            strSistema = "I"
        Case "LEGISLACAO"
            strSistema = "C"
        Case "OUVIDORIA"
            strSistema = "M"
        Case "COMPRAS"
            strSistema = "A"
        Case "PATRIMONIO"
            strSistema = "G"
        Case "MATERIAL"
            strSistema = "E"
        Case "PROTOCOLO"
            strSistema = "H"
        Case "SEGURANCA"
            strSistema = "L"
        Case "MENOR"
            strSistema = "D"
        Case "GERENCIAL"
            strSistema = "N"
    End Select
    
    'Somente diferente de 0 (zero), quando a function for utilizada para retornar um codigo sugerido  para mensagens de ]
    'duplicacao de registro
    If bitAutoNumeracaoParaMsgCriticas = 0 Then
    
        strSql = "SELECT bitAutoNumeracao FROM " & gstrItens
        strSql = strSql & " WHERE intCodigo = " & intCodigo & " AND "
        strSql = strSql & "UPPER(" & strSUBSTRING & "(strCodItem,1,1)) = '" & strSistema & "' "
            
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                        
            If adoResultado.RecordCount > 0 Then
                intOrdenacao = Abs(adoResultado.Fields("bitAutoNumeracao").Value)
            End If
        
            adoResultado.Close
                
        End If
    
    Else
    
        intOrdenacao = bitAutoNumeracaoParaMsgCriticas
        
    End If
    
    If strGrupo = "" And intOrdenacao = 2 Then intOrdenacao = 1
    
    Select Case intOrdenacao
        Case 1
'            strSQL = "SELECT MAX (convert(integer," & (strCampo) & ")) as ProximoCodigo FROM " & strTabela
            
            
            strSql = "SELECT " & gstrTOPnSQLServer(1) & " (" & gstrREPLICATE(strCampo, "0", 10) & ") as ProximoCodigo,10 - " & strLen & "(" & strCampo & ") As TotalZeros FROM " & strTabela
            If strParametroEspecifico <> "" Then
                strSql = strSql & " WHERE " & strParametroEspecifico & "=" & IIf(UCase$(Left(strParametroEspecifico, 3)) = "STR", "'" & strValorParametroEspecifico & "'", Val(strValorParametroEspecifico))
            End If
            strSql = strSql & " GROUP BY " & strCampo & " ORDER BY ProximoCodigo DESC "
            strSql = gstrTOPnOracle(strSql, 1)
            If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                
                Select Case intMascaraEspecifica
                    Case 1 'Elemento Despesa
                        If Retorno Then
                            gstrProximoCodigo = Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros"))
                            gstrProximoCodigo = gstrValorSemMascara(gvntFormatacaoEspecifica(txtDestino)) + 1
                        Else
                            txtDestino = Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros"))
                            txtDestino = gstrValorSemMascara(gvntFormatacaoEspecifica(txtDestino)) + 1
                        End If
                        adoResultado.Close
                        Set adoResultado = Nothing
                        
                    Case Else
                        If Not adoResultado.EOF Then
                            If Not IsNull(adoResultado("ProximoCodigo")) Then
                                If Retorno Then
                                    If IsNumeric(Right(Trim(adoResultado("ProximoCodigo")), Len(adoResultado("ProximoCodigo")) - adoResultado("TotalZeros"))) Then
                                        gstrProximoCodigo = (Trim(adoResultado("ProximoCodigo")) + 1)
                                        For intFor = 1 To (Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros") - Len(gstrProximoCodigo))
                                            gstrProximoCodigo = "0" & gstrProximoCodigo
                                        Next
                                    Else
                                        If IsNumeric(Right(Right(adoResultado("ProximoCodigo"), Len(adoResultado("ProximoCodigo")) - adoResultado("TotalZeros")), 1)) Then
                                            gstrProximoCodigo = Left(Right(adoResultado("ProximoCodigo"), Len(adoResultado("ProximoCodigo")) - adoResultado("TotalZeros")), Len(Right(adoResultado("ProximoCodigo"), Len(adoResultado("ProximoCodigo")) - adoResultado("TotalZeros"))) - 1) & (Right(Right(adoResultado("ProximoCodigo"), Len(adoResultado("ProximoCodigo")) - adoResultado("TotalZeros")), 1) + 1)
                                        Else
                                            gstrProximoCodigo = Left(Right(adoResultado("ProximoCodigo"), Len(adoResultado("ProximoCodigo")) - adoResultado("TotalZeros")), Len(Right(adoResultado("ProximoCodigo"), Len(adoResultado("ProximoCodigo")) - adoResultado("TotalZeros"))) - 1) & Chr(Asc(Right(Right(adoResultado("ProximoCodigo"), Len(adoResultado("ProximoCodigo")) - adoResultado("TotalZeros")), 1)) + 1)
                                        End If
                                    End If
                                Else
                                    If IsNumeric(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros"))) Then
                                        txtDestino = (Trim(adoResultado("ProximoCodigo")) + 1)
                                        For intFor = 1 To (Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros") - Len(txtDestino))
                                            txtDestino = "0" & txtDestino
                                        Next
                                    Else
                                        If IsNumeric(Right(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")), 1)) Then
                                            txtDestino = Left(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")), Len(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros"))) - 1) & (Right(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")), 1) + 1)
                                        Else
                                            txtDestino = Left(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")), Len(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros"))) - 1) & Chr(Asc(Right(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")), 1)) + 1)
                                        End If
                                    End If
                                End If
                            Else
                                If Retorno Then
                                    gstrProximoCodigo = "1"
                                Else
                                    txtDestino = "1"
                                End If
                            End If
                        Else
                            If Retorno Then
                                gstrProximoCodigo = "1"
                            Else
                                txtDestino = "1"
                            End If
                        End If
                End Select
             
                
                adoResultado.Close
                Set adoResultado = Nothing

            End If
            
        Case 2
            
            strSql = "SELECT " & gstrTOPnSQLServer(1) & " (" & gstrREPLICATE(strCampo, "0", 10) & ") as ProximoCodigo,10 - " & strLen & "(" & strCampo & ") As TotalZeros FROM " & strTabela
            strSql = strSql & " WHERE " & strGrupo & " = " & Val(strValorGrupo)
            If strSubGrupo <> "" Then
                strSql = strSql & " AND " & strSubGrupo & " = " & Val(strValorSubGrupo)
            End If
            If strParametroEspecifico <> "" Then
                strSql = strSql & " AND " & strParametroEspecifico & "=" & IIf(UCase$(Left(strParametroEspecifico, 3)) = "STR", "'" & strValorParametroEspecifico & "'", Val(strValorParametroEspecifico))
            End If
            strSql = strSql & " GROUP BY " & strCampo & " ORDER BY ProximoCodigo DESC "
            strSql = gstrTOPnOracle(strSql, 1)
            
            If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                If Not adoResultado.EOF Then
                    If Not IsNull(adoResultado("ProximoCodigo")) Then
                        If Retorno Then
                            If IsNumeric(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros"))) Then
                                gstrProximoCodigo = Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")) + 1
                            Else
                                If IsNumeric(Right(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")), 1)) Then
                                    gstrProximoCodigo = Left(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")), Len(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros"))) - 1) & (Right(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")), 1) + 1)
                                Else
                                    gstrProximoCodigo = Left(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")), Len(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros"))) - 1) & Chr(Asc(Right(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")), 1)) + 1)
                                End If
                            End If
                        Else
                            If IsNumeric(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros"))) Then
                                txtDestino = Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")) + 1
                            Else
                                If IsNumeric(Right(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")), 1)) Then
                                    txtDestino = Left(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")), Len(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros"))) - 1) & (Right(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")), 1) + 1)
                                Else
                                    txtDestino = Left(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")), Len(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros"))) - 1) & Chr(Asc(Right(Right(Trim(adoResultado("ProximoCodigo")), Len(Trim(adoResultado("ProximoCodigo"))) - adoResultado("TotalZeros")), 1)) + 1)
                                End If
                            End If
                        End If
                    Else
                        If Retorno Then
                            gstrProximoCodigo = "1"
                        Else
                            txtDestino = "1"
                        End If
                    End If
                Else
                    If strValorGrupo <> "" And strValorGrupo <> "0" Then
                        If Retorno Then
                            gstrProximoCodigo = "1"
                        Else
                            txtDestino = "1"
                        End If
                    End If
                End If
                adoResultado.Close
                Set adoResultado = Nothing
            End If
            
        Case Else
            If Retorno Then
                gstrProximoCodigo = ""
            Else
                txtDestino = ""
            End If
    End Select
    
    Set gobjBanco = Nothing
    
End Function

Public Function gblnExisteCodigo(intOrdenacao As Byte, _
                                strTabela As String, _
                                strCampo As String, _
                                strCodigo As String, _
                                Optional strGrupo As String, _
                                Optional strValorGrupo As String, _
                                Optional strSubGrupo As String, _
                                Optional strValorSubGrupo As String, _
                                Optional strWhereAdicinal As String) As Boolean

Dim strSql          As String
Dim adoResultado    As ADODB.Recordset
    
    gblnExisteCodigo = True
    
    Set gobjBanco = New clsBanco
    
    Select Case intOrdenacao
        Case 1
            
            strSql = "SELECT UPPER(" & strCampo & ") FROM " & strTabela
            strSql = strSql & " WHERE UPPER(" & strCampo & ") = "
            
            If Left(strCodigo, 1) = "'" Then
                strSql = strSql & UCase$(strCodigo)
            ElseIf UCase$(Left(strCampo, 3)) = "STR" Then
                strSql = strSql & "'" & UCase$(Trim(strCodigo)) & "'"
            ElseIf UCase$(Left(strCampo, 3)) = "INT" Or UCase$(Left(strCampo, 4)) = "PKID" Or UCase$(Left(strCampo, 3)) = "BIT" Then
                strSql = strSql & Val(strCodigo)
            End If
            
            'ACRESCENTA CLAUSULA WHERE ADICIONAL OBS. COMEÇAR COM AND
            If Trim(strWhereAdicinal) <> "" Then
                strSql = strSql & " " & Trim(strWhereAdicinal)
            End If
            
            If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                
                gblnExisteCodigo = Not adoResultado.EOF
                
                adoResultado.Close
                Set adoResultado = Nothing
            
            End If
            
        Case 2
            
            strSql = "SELECT " & strCampo & " FROM " & strTabela
            strSql = strSql & " WHERE UPPER(RTRIM(LTRIM(" & strGrupo & "))) = "
            
            If Left(strValorGrupo, 1) = "'" Then
                strSql = strSql & UCase(strValorGrupo)
            ElseIf UCase$(Left(strGrupo, 3)) = "STR" Then
                strSql = strSql & "'" & UCase(strValorGrupo) & "'"
            ElseIf UCase$(Left(strGrupo, 3)) = "INT" Or UCase$(Left(strGrupo, 4)) = "PKID" Or UCase$(Left(strGrupo, 3)) = "BYT" Or UCase$(Left(strGrupo, 3)) = "BIT" Then
                strSql = strSql & Val(strValorGrupo)
            End If
 
            If UCase$(Left(strCampo, 3)) = "DTM" Then 'Nino
                strSql = strSql & " AND " & strCampo & " = "
            Else
                strSql = strSql & " AND UPPER(" & strCampo & ") = "
            End If
            
            If Left(strCodigo, 1) = "'" Then
                strSql = strSql & UCase$(strCodigo)
            ElseIf UCase$(Left(strCampo, 3)) = "STR" Then
                strSql = strSql & "'" & UCase(strCodigo) & "'"
            ElseIf UCase$(Left(strCampo, 3)) = "INT" Or UCase$(Left(strGrupo, 4)) = "PKID" Or UCase$(Left(strCampo, 3)) = "BIT" Then
                strSql = strSql & Val(strCodigo)
            ElseIf UCase$(Left(strCampo, 3)) = "DTM" Then 'Saaalsis
            If bytDBType = EDatabases.SQLServer Then
                    strSql = strSql & CDate(strCodigo)
                Else
                    strSql = strSql & "'" & strCodigo & "'"
                End If
            End If
            
            If strSubGrupo <> "" Then
                strSql = strSql & " AND UPPER(" & strSubGrupo & ") = "
                If Left(strSubGrupo, 1) = "'" Then
                    strSql = strSql & UCase$(strValorSubGrupo)
                ElseIf UCase$(Left(strSubGrupo, 3)) = "STR" Then
                    strSql = strSql & "'" & UCase$(strValorSubGrupo) & "'"
                ElseIf UCase$(Left(strSubGrupo, 3)) = "INT" Or UCase$(Left(strGrupo, 4)) = "PKID" Or UCase$(Left(strSubGrupo, 3)) = "BIT" Then
                    strSql = strSql & Val(strValorSubGrupo)
                End If
            End If
            
            'ACRESCENTA CLAUSULA WHERE ADICIONAL OBS. COMEÇAR COM AND
            If Trim(strWhereAdicinal) <> "" Then
                strSql = strSql & " " & Trim(strWhereAdicinal)
            End If
            
            If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                
                gblnExisteCodigo = Not adoResultado.EOF
                
                adoResultado.Close
                Set adoResultado = Nothing
            
            End If
            
    End Select
    
    Set gobjBanco = Nothing

End Function

Public Function CampoObrigatorio(ByRef CaixaDeTexto As TextBox, _
                             ByVal strDescricao As String) As Boolean
                             
    If CaixaDeTexto.Text = "" Then
        ExibeMensagem "O campo " & strDescricao & " deve ser preenchido."
        CaixaDeTexto.SetFocus
        CampoObrigatorio = False
    Else
        CampoObrigatorio = True
    End If
    
End Function

Public Sub PopUpMenuMascara(frmForm As Form, Button As Integer)

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Deslocamento dos JOINS utilizados na cláusula FROM para a cláusula WHERE.
'            Foi mantida a forma antiga para o SQL Server pois não era possível o
'            deslocamento completo devido à incompatibilidade entres os bancos.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim pt              As POINTAPI
    Dim lngReturn       As Long
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    Dim adoMascara      As ADODB.Recordset
    Dim intChecado      As Integer
    
    If gblnAdmin Or gblnMaster Then
    
        lngMenu = CreatePopupMenu()
          
        strSql = "SELECT  " & gstrMascaras & ".PKId, " & gstrMascaras & ".strMascara, " & gstrMascaras & ".strDescricao, " & gstrMascarasItens & ".strCampo, " & gstrItens & ".intCodigo"
                    
        If (bytDBType = EDatabases.SQLServer) Then
            strSql = strSql & " FROM " & gstrItens & " INNER JOIN "
            strSql = strSql & gstrMascarasItens & " ON " & gstrItens & ".intCodigo = " & gstrMascarasItens & ".intItem RIGHT OUTER JOIN "
            strSql = strSql & gstrMascaras & " ON " & gstrMascarasItens & ".intItem = " & gstrMascaras & ".PKId"
        
        ElseIf (bytDBType = EDatabases.Oracle) Then
            strSql = strSql & " FROM " & gstrItens & ", "
            strSql = strSql & gstrMascarasItens & ", "
            strSql = strSql & gstrMascaras
        
            strSql = strSql & " WHERE " & gstrItens & ".intCodigo = " & gstrMascarasItens & ".intItem AND "
            strSql = strSql & gstrMascarasItens & ".intItem " & strOUTJOracle & "= " & gstrMascaras & ".PKId"
        
        End If
                
        Set gobjBanco = New clsBanco
                
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                    
            While Not adoResultado.EOF
                
                strSql = "SELECT " & gstrMascarasItens & ".strCampo "
'                strSQL = strSQL & "FROM " & gstrItens & " INNER JOIN "
                strSql = strSql & "FROM " & gstrItens & ", "
'                strSQL = strSQL & gstrMascarasItens & " ON " & gstrItens & ".intCodigo = " & gstrMascarasItens & ".intItem LEFT OUTER JOIN "
                strSql = strSql & gstrMascarasItens & ", "
'                strSQL = strSQL & gstrMascaras & " ON " & gstrMascarasItens & ".intItem = " & gstrMascaras & ".PKId"
                strSql = strSql & gstrMascaras
                strSql = strSql & " WHERE (" & gstrItens & ".intCodigo = " & gintCodSeguranca & " AND " & gstrMascarasItens & ".strCampo = '" & frmForm.ActiveControl.Name & "' and intMascara =" & adoResultado("PKId") & ") and "
    
                strSql = strSql & gstrItens & ".intCodigo = " & gstrMascarasItens & ".intItem AND "
                strSql = strSql & gstrMascarasItens & ".intItem = " & gstrMascaras & ".PKId" & strOUTJOracle
                
                Set gobjBanco = New clsBanco
                
                If gobjBanco.CriaADO(strSql, 5, adoMascara) Then
                    If Not adoMascara.EOF Then
                        AppendMenu lngMenu, MF_CHECKED, adoResultado("PKId"), CStr(adoResultado("strDescricao") & " - " & adoResultado("strMascara"))
                        intChecado = adoResultado("PKId")
                    Else
                        AppendMenu lngMenu, MF_STRING, adoResultado("PKId"), CStr(adoResultado("strDescricao") & " - " & adoResultado("strMascara"))
                    End If
                    
                End If
               
                adoResultado.MoveNext
            Wend
                    
        End If
        
        GetCursorPos pt
    
        lngReturn = TrackPopupMenuEx(lngMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, pt.X, pt.Y, frmForm.hWnd, ByVal 0&)
        
            Select Case lngReturn
                Case 0
                Case Else
                    gobjBanco.Execute "DELETE FROM " & _
                                            gstrMascarasItens & " WHERE intItem=" & gintCodSeguranca & " AND strCampo='" & frmForm.ActiveControl.Name & "'"
                    If lngReturn <> intChecado Then
                        gobjBanco.Execute "INSERT INTO " & _
                                                gstrMascarasItens & _
                                                    "(intItem, intMascara, strCampo) VALUES " & _
                                                    "(" & gintCodSeguranca & ", " & lngReturn & ", '" & frmForm.ActiveControl.Name & "')"
                    End If
                                                    
            End Select
        
        
        Set gobjBanco = Nothing
    
        DestroyMenu lngMenu
    End If
End Sub

Public Function AplicarMascara(CaixaDeTexto As Object)

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Deslocamento dos JOINS utilizados na cláusula FROM para a cláusula WHERE.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    Set gobjBanco = New clsBanco
                
'    strSQL = "SELECT " & gstrMascaras & ".strMascara FROM " & gstrMascaras & " INNER JOIN "
    strSql = "SELECT " & gstrMascaras & ".strMascara FROM " & gstrMascaras & ", "
'    strSQL = strSQL & gstrMascarasItens & " ON " & gstrMascaras & ".PKId = dbo.tblMascarasItens.intMascara "
    strSql = strSql & gstrMascarasItens
    strSql = strSql & " WHERE intItem=" & gintCodSeguranca & " AND strCampo='" & CaixaDeTexto.Name & "'"
    
    strSql = strSql & " AND " & gstrMascaras & ".PKId = tblMascarasItens.intMascara "
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            CaixaDeTexto.Text = Format(CaixaDeTexto.Text, adoResultado("strMascara"))
        End If
    End If
    
    Set gobjBanco = Nothing
                
End Function

Public Function gstrREPLICATE(strString As String, strCaracter As String, intQuantidade As Integer) As String

'******************************************************************************************
' Data: 28/05/2003
' Descrição: - strString --> String que recebera a concatenacao.
'            - strCaracter --> Caracter a ser duplicado
'            - intQuantidade --> Qauntidade de vezes que o caracter será duplicado na string
' Responsável: Gustavo Monteiro
'******************************************************************************************

    If (bytDBType = EDatabases.Oracle) Then
        gstrREPLICATE = " lPad(" & strString & ", " & intQuantidade & ",'" & strCaracter & "')"
    Else
        gstrREPLICATE = " REPLICATE('" & strCaracter & "', " & intQuantidade & " - Len(" & strString & ")) + " & gstrCONVERT(CDT_VARCHAR, strString)
    End If

End Function

'Public Function gblnDadosOk(frmForm As Form) As Boolean
'    Dim objControl          As Object
'    Dim objDescricao        As Object
'    Dim strLabel            As String
'    Dim intFor              As Integer
'
'    For intFor = frmForm.Controls.Count - 1 To 0 Step -1
'        If Mid(frmForm.Controls(intFor).Name, 3, 1) <> "_" Then
'            If TypeOf frmForm.Controls(intFor) Is TextBox Then
'                If frmForm.Controls(intFor).Enabled Then
'                    If Not frmForm.Controls(intFor).DragMode = 1 And frmForm.Controls(intFor).Text = "" Then
'                        On Error Resume Next
'                        strLabel = frmForm.Controls(Replace(frmForm.Controls(intFor).Name, "txt", "lbl")).Caption
'
'                        If Err.Number = 0 Then
'                            ExibeMensagem "Preencha corretamente o campo " & strLabel & "!", vbInformation
'                        Else
'                            ExibeMensagem "Este campo dever ser preenchido!", vbInformation
'                        End If
'
'                        frmForm.Controls(intFor).SetFocus
'
'                        Exit Function
'
'                        On Error GoTo 0
'                    End If
'                End If
'            ElseIf TypeOf frmForm.Controls(intFor) Is DataCombo Then
'                If frmForm.Controls(intFor).Enabled Then
'                    If Not frmForm.Controls(intFor).DragMode = 1 And Not frmForm.Controls(intFor).MatchedWithList Then
'
'                        On Error Resume Next
'                        strLabel = frmForm.Controls(Replace(frmForm.Controls(intFor).Name, "dbc", "lbl")).Caption
'
'                        If Err.Number = 0 Then
'                            ExibeMensagem "Preencha corretamente o campo " & strLabel, vbInformation
'                        Else
'                            ExibeMensagem "Este campo dever ser preenchido!", vbInformation
'                        End If
'
'                        On Error GoTo 0
'
'                        frmForm.Controls(intFor).SetFocus
'
'                        Exit Function
'                    End If
'                End If
'            ElseIf TypeOf frmForm.Controls(intFor) Is ComboBox Then
'                If frmForm.Controls(intFor).Enabled Then
'                    If Not frmForm.Controls(intFor).DragMode = 1 And frmForm.Controls(intFor).Text = "" Then
'
'                        On Error Resume Next
'
'                        strLabel = frmForm.Controls(Replace(frmForm.Controls(intFor).Name, "cbo", "lbl")).Caption
'
'                        If Err.Number = 0 Then
'                            ExibeMensagem "Preencha corretamente o campo " & strLabel, vbInformation
'                        Else
'                            ExibeMensagem "Este campo dever ser preenchido!", vbInformation
'                        End If
'
'                        On Error GoTo 0
'
'                        frmForm.Controls(intFor).SetFocus
'
'                        Exit Function
'                    End If
'                End If
'            End If
'        End If
'    Next
'
'    gblnDadosOk = True
'End Function

Public Sub LeModuloAtual()

    Dim adoResultado    As ADODB.Recordset
    Dim strSistema      As String
    
    gstrContribuinteTituloPl = "Contribuinte"
    gstrContribuinteTituloSg = "Contribuinte"
    
    Select Case UCase(App.ProductName)
        Case "TRIBUTARIO"
            strSistema = "J"
            gstrContribuinteTituloPl = "Contribuintes"
            gstrContribuinteTituloSg = "Contribuinte"
        Case "ORCAMENTARIO"
            strSistema = "F"
            gstrContribuinteTituloPl = "Credores"
            gstrContribuinteTituloSg = "Credor"
        Case "FROTA"
            strSistema = "B"
        Case "RH"
            strSistema = "I"
        Case "LEGISLACAO"
            strSistema = "C"
        Case "OUVIDORIA"
            strSistema = "M"
            gstrContribuinteTituloPl = "Solicitantes"
            gstrContribuinteTituloSg = "Solicitante"
        Case "COMPRAS"
            strSistema = "A"
            gstrContribuinteTituloPl = "Proponentes"
            gstrContribuinteTituloSg = "Proponente"
        Case "PATRIMONIO"
            strSistema = "G"
            gstrContribuinteTituloPl = "Fornecedores"
            gstrContribuinteTituloSg = "Fornecedor"
        Case "MATERIAL"
            strSistema = "E"
            gstrContribuinteTituloPl = "Fornecedores"
            gstrContribuinteTituloSg = "Fornecedor"
        Case "PROTOCOLO"
            strSistema = "H"
            gstrContribuinteTituloPl = "Requerentes"
            gstrContribuinteTituloSg = "Requerente"
        Case "SEGURANCA"
            strSistema = "L"
        Case "MENOR"
            strSistema = "D"
        Case "GERENCIAL"
            strSistema = "N"
    End Select

    Set gobjBanco = New clsBanco

    If gobjBanco.CriaADO("SELECT PKId FROM " & gstrItens & " WHERE strCodItem = '" & strSistema & "'", 5, adoResultado) Then
        gintModulo = adoResultado(0)
    End If

    Set gobjBanco = Nothing
    
End Sub

Public Sub CepLogradouro(strCep As String, _
                            objEndereco As Object, _
                            Optional objBairro As Object, _
                            Optional objCidade As Object, _
                            Optional objEstado As Object, _
                            Optional objTipoLogradouro As Object, _
                            Optional objTituloLogradouro As Object, _
                            Optional ByRef blnPertenceAoMunicipio As Boolean, _
                            Optional blnBoundTextEndereco As Boolean, _
                            Optional blnBoundTextBairro As Boolean = False, _
                            Optional blnBoundTextMunicipio As Boolean = False, _
                            Optional blnBoundTextUF As Boolean = False, _
                            Optional blnBoundTextTipoLogradouro As Boolean = False, _
                            Optional blnBoundTextTituloLogradouro As Boolean = False, _
                            Optional blnConcatenarTipoTitulo As Boolean = False, _
                            Optional blnSegundaPesquisa As Boolean = True, _
                            Optional strParametros As String, _
                            Optional blnMostraLogradourosCancelados As Boolean = False)
                            
    Dim adoResultado        As ADODB.Recordset
    Dim strSql              As String
    Dim cmdBotao            As CommandButton
    Dim strTag              As String
    
    Set gobjBanco = New clsBanco

    If strCep = "" Then
        Exit Sub
    End If

    If objBairro Is Nothing Then
        strSql = "SELECT " & IIf(blnBoundTextEndereco, "PKId, ", "") & "strDescricao FROM " & gstrLogradouro
        strSql = strSql & " WHERE intCep =" & Replace(strCep, "-", "")
        strSql = strSql & " AND Dtmdtexclusao is null "
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                If blnBoundTextEndereco Then
                    objEndereco.BoundText = adoResultado.Fields("PKId")
                End If
                objEndereco.Text = LTrim(RTrim(adoResultado.Fields("strLogradouro")))
            Else
                If blnSegundaPesquisa Then
                    strSql = "SELECT " & IIf(blnBoundTextEndereco, "PKId, ", "") & "strLogradouro FROM " & gstrCeps
                    strSql = strSql & " WHERE intCep =" & Replace(strCep, "-", "")
                    strSql = strSql & " AND Dtmdtexclusao is null "
                
                    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                            
                        If Not adoResultado.EOF Then
                            If blnBoundTextEndereco Then
                                objEndereco.BoundText = adoResultado.Fields("PKId")
                            End If
                            objEndereco.Text = LTrim(RTrim(adoResultado.Fields("strDescricao")))
                        Else
                            objEndereco.Text = ""
                        End If
                    End If
                End If
            End If
        End If
    Else
        If (bytDBType = EDatabases.Oracle) Then
            
            strSql = "SELECT "
            strSql = strSql & "CE.PKId AS intLogradouro, CE.strDescricao AS strLogradouro, CE.intTipoLogradouro as intTipo, CE.intTituloLogradouro as intTitulo, "
            strSql = strSql & "BA.PKId AS intBairro, BA.bytPertenceAoMunicipo As PertenceAoMunicipio, BA.strDescricao AS strBairro "
            strSql = strSql & ",(SELECT MU.strDescricao FROM " & gstrCidade & " MU WHERE BA.intMunicipio = MU.PKId) AS strMunicipio "
            strSql = strSql & ",(SELECT UF.strSigla FROM " & gstrUF & " UF WHERE UF.PKId = (SELECT MU.intUF FROM " & gstrCidade & " MU WHERE BA.intMunicipio = MU.PKId )) AS strEstado "
            strSql = strSql & ",(SELECT TL.strDescricao FROM " & gstrTipoLogradouro & " TL WHERE CE.intTipoLogradouro = TL.PKId) AS strTipo "
            strSql = strSql & ",(SELECT TL.strSigla FROM " & gstrTituloLogradouro & " TL WHERE CE.intTituloLogradouro = TL.PKId) AS strTitulo "
    
            strSql = strSql & ",(SELECT MU.PKId FROM " & gstrCidade & " MU WHERE BA.intMunicipio = MU.PKId) AS intMunicipio "
            strSql = strSql & ",(SELECT UF.PKId FROM " & gstrUF & " UF WHERE UF.PKId = (SELECT MU.intUF FROM " & gstrCidade & " MU WHERE BA.intMunicipio = MU.PKId )) AS intEstado "
            
            strSql = strSql & "FROM "
            strSql = strSql & gstrLogradouro & " CE, "
            strSql = strSql & gstrBairro & " BA "
            strSql = strSql & "WHERE "
            strSql = strSql & "CE.intBairro = BA.PKId AND "
            strSql = strSql & "CE.intCep = " & Replace(strCep, "-", "")
            
            If Not blnMostraLogradourosCancelados Then
                strSql = strSql & " AND CE.Dtmdtexclusao IS NULL "
            End If

        Else
            strSql = "SELECT "
            strSql = strSql & "CE.PKId AS intLogradouro, CE.strDescricao AS strLogradouro, CE.intTipoLogradouro as intTipo, CE.intTituloLogradouro as intTitulo, "
            strSql = strSql & "BA.PKId AS intBairro, BA.bytPertenceAoMunicipo As PertenceAoMunicipio, BA.strDescricao AS strBairro "
            strSql = strSql & ",(SELECT MU.strDescricao FROM " & gstrCidade & " MU WHERE BA.intMunicipio = MU.PKId) AS strMunicipio "
            strSql = strSql & ",(SELECT UF.strSigla FROM " & gstrUF & " UF WHERE UF.PKId = (SELECT MU.intUF FROM " & gstrCidade & " MU WHERE BA.intMunicipio = MU.PKId )) AS strEstado "
            strSql = strSql & ",(SELECT TL.strDescricao FROM " & gstrTipoLogradouro & " TL WHERE CE.intTipoLogradouro = TL.PKId) AS strTipo "
            strSql = strSql & ",(SELECT TL.strSigla FROM " & gstrTituloLogradouro & " TL WHERE CE.intTituloLogradouro = TL.PKId) AS strTitulo "
    
            strSql = strSql & ",(SELECT MU.PKId FROM " & gstrCidade & " MU WHERE BA.intMunicipio = MU.PKId) AS intMunicipio "
            strSql = strSql & ",(SELECT UF.PKId FROM " & gstrUF & " UF WHERE UF.PKId = (SELECT MU.intUF FROM " & gstrCidade & " MU WHERE BA.intMunicipio = MU.PKId )) AS intEstado "
            
            strSql = strSql & "FROM "
            strSql = strSql & gstrLogradouro & " CE, "
            strSql = strSql & gstrBairro & " BA "
            strSql = strSql & "WHERE "
            strSql = strSql & "CE.intBairro = BA.PKId AND "
            strSql = strSql & "CE.intCep = " & Replace(strCep, "-", "")
            
            If Not blnMostraLogradourosCancelados Then
                strSql = strSql & " AND CE.Dtmdtexclusao IS NULL "
            End If
            
        End If
        
            If strParametros <> "" Then
                strSql = strSql & " AND " & strParametros
            End If
        
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                If blnConcatenarTipoTitulo Then
                    If blnBoundTextEndereco Then
                        If TypeOf objEndereco Is DataCombo Then AtribuiValorDoSql objEndereco, adoResultado.Fields("intLogradouro")
                        If adoResultado.RecordCount > 1 Then
                            Load frmSelecionaLogradouro
                            frmSelecionaLogradouro.strQuery = strSql
                            
                            frmSelecionaLogradouro.Caption = "Selecione o logradouro desejado para o CEP " & strCep
                            
                            LeDaTabelaParaObj gstrCeps, frmSelecionaLogradouro.tdb_Lista, strSql
                            
                            CarregaForm frmSelecionaLogradouro, objEndereco, strSql, vbModal
                            
                            adoResultado.Filter = "intLogradouro =" & objEndereco.BoundText
                            
                            PreencherListaDeOpcoes objEndereco, objEndereco.BoundText
                        Else
                            PreencherListaDeOpcoes objEndereco, adoResultado.Fields("intLogradouro")
                        End If
                    Else
                        If adoResultado.RecordCount > 1 Then
                            
                            Load frmSelecionaLogradouro
                            frmSelecionaLogradouro.strQuery = strSql
                            
                            frmSelecionaLogradouro.Caption = "Selecione o logradouro desejado para o CEP " & strCep
                            
                            LeDaTabelaParaObj gstrCeps, frmSelecionaLogradouro.tdb_Lista, strSql
                            
                            CarregaForm frmSelecionaLogradouro, objEndereco, strSql, vbModal
                            
                            If Trim(objEndereco.Tag) <> "" Then
                                adoResultado.Filter = "intLogradouro =" & objEndereco.Tag
                            End If
                            
                            objEndereco.Text = LTrim(RTrim(adoResultado.Fields("strLogradouro"))) & ""
                        Else
                            objEndereco.Text = LTrim(RTrim(adoResultado.Fields("strLogradouro"))) & ""
                        End If
                    End If
                    'objEndereco.Text = IIf(IsNull(adoResultado.Fields("strTitulo")), "", adoResultado.Fields("strTitulo") & " ") & LTrim(RTrim(adoResultado.Fields("strLogradouro"))) & IIf(IsNull(adoResultado.Fields("strTipo")), "", ", ") & adoResultado.Fields("strTipo")
                Else
                    If blnBoundTextEndereco Then
                        If TypeOf objEndereco Is DataCombo Then AtribuiValorDoSql objEndereco, adoResultado.Fields("intLogradouro")
                        If adoResultado.RecordCount > 1 Then
                            
                            Load frmSelecionaLogradouro
                            frmSelecionaLogradouro.strQuery = strSql
                            
                            frmSelecionaLogradouro.Caption = "Selecione o logradouro desejado para o CEP " & strCep
                            
                            LeDaTabelaParaObj gstrCeps, frmSelecionaLogradouro.tdb_Lista, strSql
                            
                            CarregaForm frmSelecionaLogradouro, objEndereco, strSql, vbModal
                            
                            adoResultado.Filter = "intLogradouro =" & objEndereco.BoundText
                            
                            PreencherListaDeOpcoes objEndereco, objEndereco.BoundText
                        Else
                            PreencherListaDeOpcoes objEndereco, adoResultado.Fields("intLogradouro")
                        End If
                    Else
                        If adoResultado.RecordCount > 1 Then
                            
                            Load frmSelecionaLogradouro
                            frmSelecionaLogradouro.strQuery = strSql
                            
                            frmSelecionaLogradouro.Caption = "Selecione o logradouro desejado para o CEP " & strCep
                            
                            LeDaTabelaParaObj gstrCeps, frmSelecionaLogradouro.tdb_Lista, strSql
                            
                            CarregaForm frmSelecionaLogradouro, objEndereco, strSql, vbModal
                            If Trim(objEndereco.Tag) <> "" Then
                                adoResultado.Filter = "intLogradouro =" & objEndereco.Tag
                            End If
                            objEndereco.Text = LTrim(RTrim(adoResultado.Fields("strLogradouro"))) & ""
                        Else
                            objEndereco.Text = LTrim(RTrim(adoResultado.Fields("strLogradouro"))) & ""
                        End If
                        
                    End If
                    
                    
                    If Not objTipoLogradouro Is Nothing Then
                        objTipoLogradouro.Text = LTrim(RTrim(adoResultado.Fields("strTipo"))) & ""
                        If blnBoundTextTipoLogradouro Then
                            objTipoLogradouro.BoundText = adoResultado.Fields("intTipo") & ""
                            If Not IsNull(adoResultado.Fields("intTipo")) Then PreencherListaDeOpcoes objTipoLogradouro, adoResultado.Fields("intTipo")
                        'Else
                        '    objTipoLogradouro.Tag = adoResultado.Fields("intTipo") & ""
                        End If
                    End If
                    
                    If Not objTituloLogradouro Is Nothing Then
                        objTituloLogradouro.Text = LTrim(RTrim(adoResultado.Fields("strTitulo"))) & ""
                        If blnBoundTextTituloLogradouro Then
                            objTituloLogradouro.BoundText = adoResultado.Fields("intTitulo") & ""
                            If Not IsNull(adoResultado.Fields("intTitulo")) Then PreencherListaDeOpcoes objTituloLogradouro, adoResultado.Fields("intTitulo")
                        'Else
                        '    objTituloLogradouro.Tag = adoResultado.Fields("intTitulo") & ""
                        End If
                    End If
                End If
                
                
                    
                If blnBoundTextBairro Then
                    objBairro.BoundText = adoResultado.Fields("intbairro") & ""
                    PreencherListaDeOpcoes objBairro, adoResultado.Fields("intBairro")
                'Else
                '    objBairro.Tag = adoResultado.Fields("intbairro") & ""
                End If
                objBairro.Text = LTrim(RTrim(adoResultado.Fields("strBairro"))) & ""
                    
                If Not objCidade Is Nothing Then
                    objCidade.Text = LTrim(RTrim(adoResultado.Fields("strMunicipio"))) & ""
                    If blnBoundTextMunicipio Then
                        objCidade.BoundText = adoResultado.Fields("intMunicipio") & ""
                        PreencherListaDeOpcoes objCidade, adoResultado.Fields("intMunicipio")
                    'Else
                    '    objCidade.Tag = adoResultado.Fields("intMunicipio") & ""
                    End If
                End If
                    
                If Not objEstado Is Nothing Then
                    objEstado.Text = LTrim(RTrim(adoResultado.Fields("strEstado"))) & ""
                    If blnBoundTextUF Then
                        objEstado.BoundText = adoResultado.Fields("intEstado") & ""
                        PreencherListaDeOpcoes objEstado, adoResultado.Fields("intEstado")
                    'Else
                    '    objEstado.Tag = adoResultado.Fields("intEstado") & ""
                    End If
                End If
                
                blnPertenceAoMunicipio = IIf(IsNull(adoResultado.Fields("PertenceAoMunicipio")), 0, adoResultado.Fields("PertenceAoMunicipio"))
                    
            Else
                If blnSegundaPesquisa Then
                    'Segunda pesquisa
                    strSql = "SELECT CE.PKID AS intLogradouro, "
                    strSql = strSql & "CE.strLogradouro AS strLogradouro, CE.strUF AS strEstado, "
                    strSql = strSql & "CE.strBairro AS strBairro,CE.strMunicipio AS strMunicipio, "
                    strSql = strSql & "MU.Pkid AS intMunicipio, UF.Pkid AS intUF "
                    strSql = strSql & "FROM "
                    strSql = strSql & gstrCeps & " CE, "
                    strSql = strSql & gstrCidade & " MU, "
                    strSql = strSql & gstrUF & " UF "
                    strSql = strSql & "WHERE "
                    strSql = strSql & "CE.intCep = " & Replace(strCep, "-", "")
                    strSql = strSql & " AND LTrim(Upper(CE.strMunicipio)) " & strOUTJSQLServer & "= LTrim(Upper(MU.strDescricao " & strOUTJOracle & "))"
                    strSql = strSql & " AND LTrim(Upper(CE.strUF)) " & strOUTJSQLServer & "= LTrim(Upper(UF.strSigla " & strOUTJOracle & "))"
                    
                    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                        If Not adoResultado.EOF Then
                                
                            If adoResultado.RecordCount > 1 Then
                                Load frmSelecionaLogradouro
                                frmSelecionaLogradouro.strQuery = strSql
                                frmSelecionaLogradouro.Caption = "Selecione o logradouro desejado para o CEP " & strCep
                                LeDaTabelaParaObj gstrCeps, frmSelecionaLogradouro.tdb_Lista, strSql
                                CarregaForm frmSelecionaLogradouro, objEndereco, strSql, vbModal
                            
                                If blnBoundTextEndereco Then
                                    adoResultado.Filter = "intLogradouro =" & objEndereco.BoundText
                                Else
                                    adoResultado.Filter = "intLogradouro =" & Val(objEndereco.Tag)
                                End If
                                'objEndereco.Text = LTrim(RTrim(adoResultado.Fields("strLogradouro"))) & ""
                            Else
                                objEndereco.Text = LTrim(RTrim(adoResultado.Fields("strLogradouro"))) & ""
                            End If
                            
                            objBairro.Text = LTrim(RTrim(adoResultado.Fields("strBairro"))) & ""
                                    
                            If Not objCidade Is Nothing Then
                                If blnBoundTextMunicipio Then
                                    AtribuiValorDoSql objCidade, adoResultado.Fields("intMunicipio")
                                Else
                                    objCidade.Text = LTrim(RTrim(adoResultado.Fields("strMunicipio"))) & ""
                                End If
                            End If
                                    
                            If Not objEstado Is Nothing Then
                                If blnBoundTextUF Then
                                    AtribuiValorDoSql objEstado, adoResultado.Fields("intUF")
                                Else
                                    objEstado.Text = LTrim(RTrim(adoResultado.Fields("strEstado"))) & ""
                                End If
                            End If
                                    
                        'Else
                        '    On Error Resume Next
                        '    If objEndereco.Tag = "" Then
                        '        If TypeOf objEndereco Is DataCombo Then Set objEndereco.DataSource = Nothing
                        '    End If
                        '    objEndereco.Text = ""
                        '    objBairro.Text = ""
                        '    If Not objCidade Is Nothing Then objCidade.Text = ""
                        '    If Not objEstado Is Nothing Then objEstado.Text = ""
                        '    On Error GoTo 0
                    
                        End If
                    End If
                Else
                    On Error Resume Next
                    If objEndereco.Tag = "" Then
                        If TypeOf objEndereco Is DataCombo Then Set objEndereco.DataSource = Nothing
                    End If
                    objEndereco.Text = ""
                    objBairro.Text = ""
                    If Not objCidade Is Nothing Then objCidade.Text = ""
                    If Not objEstado Is Nothing Then objEstado.Text = ""
                    On Error GoTo 0
                End If
            End If
        End If
    End If
    
    Set gobjBanco = Nothing
    
End Sub

Public Sub LogradouroCep(lngLogradouro As Long, _
                            Optional objBairro As Object, _
                            Optional blnBoundTextBairro As Boolean = False, _
                            Optional objCidade As Object, _
                            Optional objEstado As Object, _
                            Optional objCep As Object, _
                            Optional blnBoundTextMunicipio As Boolean = False, _
                            Optional blnBoundTextUF As Boolean = False, _
                            Optional blnExcluidos As Boolean = True)

    Dim adoResultado        As ADODB.Recordset
    Dim strSql              As String

    Set gobjBanco = New clsBanco

    strSql = "SELECT LO.intCep, "
    strSql = strSql & "BA.PKId AS intBairro, BA.strDescricao AS strBairro "
    If Not objCidade Is Nothing Then strSql = strSql & ",(SELECT MU.strDescricao FROM " & gstrCidade & " MU WHERE BA.intMunicipio = MU.PKId) AS strMunicipio "
    If Not objEstado Is Nothing Then strSql = strSql & ",(SELECT UF.strSigla FROM " & gstrUF & " UF WHERE UF.PKId = (SELECT MU.intUF FROM " & gstrCidade & " MU WHERE BA.intMunicipio = MU.PKId )) AS strEstado "
    If Not objCidade Is Nothing Then strSql = strSql & ",(SELECT MU.PKId FROM " & gstrCidade & " MU WHERE BA.intMunicipio = MU.PKId) AS intMunicipio "
    If Not objEstado Is Nothing Then strSql = strSql & ",(SELECT UF.PKId FROM " & gstrUF & " UF WHERE UF.PKId = (SELECT MU.intUF FROM " & gstrCidade & " MU WHERE BA.intMunicipio = MU.PKId )) AS intEstado "
    strSql = strSql & "FROM "
    strSql = strSql & gstrBairro & " BA, "
    strSql = strSql & gstrLogradouro & " LO "
    strSql = strSql & "WHERE "
    strSql = strSql & "LO.intBairro = BA.PKId AND "
    strSql = strSql & "LO.PKId = " & lngLogradouro
    If Not blnExcluidos Then
        strSql = strSql & " AND LO.Dtmdtexclusao is null "
    End If
    
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            With adoResultado

                If Not .EOF Then

                    If blnBoundTextBairro Then
                        objBairro.BoundText = .Fields("intbairro") & ""
                    End If
                    objBairro.Text = LTrim(RTrim(.Fields("strBairro"))) & ""
                    
                    If Not objCidade Is Nothing Then
                        objCidade.Text = LTrim(RTrim(.Fields("strMunicipio"))) & ""
                        If blnBoundTextMunicipio Then
                            objCidade.BoundText = .Fields("intMunicipio") & ""
                        'Else
                        '    objCidade.Tag = .Fields("intMunicipio") & ""
                        End If
                    End If
                    
                    'Nino
                    If Not objEstado Is Nothing Then
                        If TypeOf objEstado Is DataCombo Then
                            If Not objEstado <> "" Then
                                If Not objEstado.ListField <> "" Then
                                    objEstado.ListField = .Fields("strEstado").Name
                                    objEstado.BoundColumn = .Fields("intEstado").Name
                                    Set objEstado.RowSource = adoResultado
                                End If
                            End If
                        Else
                            objEstado.Text = LTrim(RTrim(.Fields("strEstado"))) & ""
                        End If
                        If blnBoundTextUF Then
                            objEstado.BoundText = .Fields("intEstado") & ""
                        End If
                    End If

                    If Not objCep Is Nothing Then objCep.Text = gstrCEPFormatado(LTrim(RTrim(.Fields("intCep"))) & "")

                Else
                    On Error Resume Next
                    objBairro.Text = ""
                    If Not objCidade Is Nothing Then objCidade.Text = ""
                    If Not objEstado Is Nothing Then objEstado.Text = ""
                    On Error GoTo 0
                End If
            End With
        End If
    Set gobjBanco = Nothing

End Sub

Public Function lngRetornaPai(lngLocal As Long) As Long
    Dim adoFilho    As ADODB.Recordset
    Dim adoPai      As ADODB.Recordset
    Dim blnOk       As Boolean
    Dim lngPai      As Long

    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO("SELECT PKId, intLocalSub FROM " & gstrLocais & " WHERE PKId = " & lngLocal, 5, adoFilho) Then
        lngPai = gstrENulo(adoFilho("intLocalSub"))
        While Not blnOk
            If gobjBanco.CriaADO("SELECT PKId, intLocalSub FROM " & gstrLocais & " WHERE PKId = " & lngPai, 5, adoPai) Then
            
                If Not adoPai.EOF Then
                    If IsNull(adoPai("intLocalSub")) Then lngRetornaPai = lngPai: Exit Function
                    If lngPai = Val(gstrENulo(adoPai("intLocalSub"))) Then
                        lngRetornaPai = lngPai
                        Exit Function
                    End If
                    lngPai = Val(gstrENulo(adoPai("intLocalSub")))
                Else
                    blnOk = True
                    lngRetornaPai = lngPai
                End If

            End If
        Wend
    End If
    
End Function

Private Sub AjustaPersonalizar(ab As ActiveBar2)
    ab.Localize ddLTCustomCaption, "Personalizar"
    ab.Localize ddLTToolbarTab, "Barras de &Ferramentas"
    ab.Localize ddLTNewButton, "&Nova"
    ab.Localize ddLTNewToolbarCaption, "Nova Barra de Ferramentas"
    ab.Localize ddLTToolbarName, "&Nome da Barra de Ferramentas"
    ab.Localize ddLTRenameButton, "&Renomear"
    ab.Localize ddLTRenameCaption, "Renomear Barra de ferramentas"
    ab.Localize ddLTDeleteButton, "&Excluir"
    ab.Localize ddLTDeleteToolbarCaption, "Excluir Barra de Ferramentas"
    ab.Localize ddLTDeleteToolString, "Tem certeza que deseja excluir a barra de ferramenta?"
    ab.Localize ddLTResetButton, "Re&definir"
    ab.Localize ddLTCommandTab, "Co&mandos"
    ab.Localize ddLTOptionsTab, "O&pções"
    ab.Localize ddLTLargeIcons, "Ícones &Grandes"
    ab.Localize ddLTScreenTips, "Mostrar &dicas de tela nas barras de ferrammentas"
    ab.Localize ddLTShortcutKeys, "Mostrar teclas de atal&ho nas dicas de tela"
    ab.Localize ddLTMenuAnimationLabel, "&Animações de Menu"
    ab.Localize ddLTMANone, "(Nenhuma)"
    ab.Localize ddLTMARandom, "Aleatórias"
    ab.Localize ddLTMASlide, "Desdobradas"
    ab.Localize ddLTMAUnfold, "Deslizantes"
    ab.Localize ddLTKeyboardButton, "&Teclado"
    ab.Localize ddLTCloseButton, "Fechar"
    ab.Localize ddLTOkButton, "Ok"
    ab.Localize ddLTCancelButton, "Cancelar"
    ab.Localize ddLTCategoriesLabel, "&Categorias"
    ab.Localize ddLTCommandLabel, "C&omandos"
    ab.Localize ddLTPressNewShortcutLabel, "Pressione &nova tecla de atalho"
    ab.Localize ddLTCurrentKeysLabel, "&Teclas Atuais"
    ab.Localize ddLTAssignButton, "At&ribuir"
    ab.Localize ddLTRemoveButton, "Re&mover"
    ab.Localize ddLTResetAllButton, "Re&definir tudo"
    ab.Localize ddLTDescription, "Descrição"
    ab.Localize ddLTMenuCustomize, "Personali&zar"
    ab.Localize ddLTKeyboardCaption, "Personalizar Teclado"
    ab.Localize ddLTMenuShowMRUFirst, "Me&nus mostram primeiro menus recém usados"
    ab.Localize ddLTShowFullMenuAfterDelay, "Mostrar menu completo após um pequeno intervalo"
    ab.Localize ddLTResetUsageData, "&Redefinir dados de uso"
    ab.Localize ddLTOther, "Outros"
    ab.Localize ddLTPersonalizedMenu, "Menus e barras de ferramentas personalizadas"
    ab.Localize ddLTCommandDesc, "Descrição"
    ab.Localize ddLTMoreButton, "Mais Botões"
    ab.Localize ddLTAlt, "Alt"
    ab.Localize ddLTControl, "Control"
    ab.Localize ddLTShift, "Shift"
    ab.Localize ddLTModifySelection, "Modificar s&eleção"
    ab.Localize ddLTMinimizeButton, "Minimizar"
    ab.Localize ddLTRestoreButton, "Restaurar"
    ab.Localize ddLTCloseWindowButton, "Fechar"
End Sub

Public Sub PulaLinha(ByVal rptRelatorio As ActiveReport, ByVal intEspacamento As Integer, Optional intAlturaDetalhe As Integer)
Dim intControle As Integer
    
    DoEvents
    
    If intAlturaDetalhe = 0 Then intAlturaDetalhe = 250
    
    With rptRelatorio
        
        .Detail.Height = intAlturaDetalhe + intEspacamento
        
        For intControle = 0 To .Detail.Controls.Count - 1
               
            .Detail.Controls.Item(intControle).Top = intEspacamento
            
        Next
        
    End With
    
End Sub

Public Sub NegritaLinha(ByVal rptRelatorio As ActiveReport, Optional blnStatus As Boolean = True)
Dim intControle As Integer
    DoEvents
    
    With rptRelatorio
        
        For intControle = 0 To .Detail.Controls.Count - 2
            
            .Detail.Controls.Item(intControle).Font.Bold = blnStatus
            
        Next
    
    End With
    
End Sub

Public Function ProcessoArquivado(lngPkidVolume As Long) As Boolean
    Dim adoTemp As ADODB.Recordset
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & " SELECT " & gstrTOPnSQLServer(1) & " pkid, bytTipo FROM " & gstrArquivamentoProcesso
    strSql = strSql & " WHERE intProtocolizacaoDoVolume = " & lngPkidVolume
    strSql = strSql & " ORDER BY pkid DESC "
    
    strSql = gstrTOPnOracle(strSql, 1)
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoTemp) Then
        If Not (adoTemp.BOF And adoTemp.EOF) Then
            ProcessoArquivado = adoTemp!bytTipo = 0
        End If
    End If
    Set gobjBanco = Nothing
    
End Function



Public Function ProtocoloJaJuntado(PkidProcesso As Long, Optional blnExibeMensagem As Boolean = True) As Boolean

'******************************************************************************************
' Data: 14/03/2003
' Alteração: - Alteração do nome do atributo intProtocolizacaoProcessoInicial.
' Responsável: Everton Bianchini
'******************************************************************************************
Dim strSql      As String
Dim adoTemp     As ADODB.Recordset
Dim strMensagem As String

        strSql = "SELECT B.strCodigo " & strCONCAT & "'-'" & strCONCAT & " LTrim(" & gstrCONVERT(CDT_VARCHAR, "B.bitDigito)") & strCONCAT & "'/'" & strCONCAT & " LTrim(" & gstrCONVERT(CDT_VARCHAR, "B.intExercicio)") & strCONCAT & "' vol.'" & strCONCAT & "  LTrim(" & gstrCONVERT(CDT_VARCHAR, "D.intVolume)") & "  As strProtocolo, "
        strSql = strSql & " B2.strCodigo " & strCONCAT & "'-'" & strCONCAT & "  LTrim(" & gstrCONVERT(CDT_VARCHAR, "B2.bitDigito)") & strCONCAT & "'/'" & strCONCAT & "  LTrim(" & gstrCONVERT(CDT_VARCHAR, "B2.intExercicio)") & strCONCAT & "' vol.'" & strCONCAT & "  LTrim(" & gstrCONVERT(CDT_VARCHAR, "D2.intVolume)") & "  As strProtocoloCapa "
        strSql = strSql & " FROM " & gstrJuntada & " A, "
        strSql = strSql & gstrProtocolizacaoProcesso & " B, "
        strSql = strSql & gstrProtocolizacaoProcesso & " B2, "
        strSql = strSql & gstrProtocolizacaoVolume & " D " & strREADPAST & ", "
        strSql = strSql & gstrProtocolizacaoVolume & " D2 " & strREADPAST
        strSql = strSql & " WHERE A.intProtocolizacaoVolumeInici = " & PkidProcesso
        strSql = strSql & " AND D.PKId = A.intProtocolizacaoVolumeInici "
        strSql = strSql & " AND D2.PKId = A.intProtocolizacaoVolume "
        strSql = strSql & " AND B.PKId = D.intProtocolizacaoProcesso "
        strSql = strSql & " AND B2.PKId = D2.intProtocolizacaoProcesso "
        strSql = strSql & " AND bitVinculado = 1"
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoTemp) Then
            If Not (adoTemp.BOF And adoTemp.EOF) Then
                ProtocoloJaJuntado = True
                If blnExibeMensagem Then ExibeMensagem " O Processo " & adoTemp!strProtocolo & " já está apensado ao " & adoTemp!strprotocolocapa & "."
            End If
        End If
        Set gobjBanco = Nothing
        
End Function

Public Function ProtocoloComApenso(PkidProcesso As Long) As Boolean
Dim strSql      As String
Dim adoTemp     As ADODB.Recordset
Dim strMensagem As String
        
        strSql = "SELECT (SELECT COUNT(bytTipoProcesso) FROM " & gstrJuntada & " WHERE bytTipoProcesso = 2 AND intProtocolizacaoVolume  = " & PkidProcesso & ")"
        strSql = strSql & " - " 'Subtracao
        strSql = strSql & " (SELECT COUNT(bytTipoProcesso) FROM " & gstrJuntada & " WHERE bytTipoProcesso = 3 AND intProtocolizacaoVolume  = " & PkidProcesso & ") Apensos"
        strSql = strSql & " FROM " & gstrJuntada & " WHERE intProtocolizacaoVolume  = " & PkidProcesso
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoTemp) Then
            If Not (adoTemp.BOF And adoTemp.EOF) Then
                ProtocoloComApenso = adoTemp!Apensos <> 0
            End If
        End If
        Set gobjBanco = Nothing
                                                        
End Function

Public Function ProtocoloEmRemessa(PkidProcesso As Long, Optional blnExibeMensagem As Boolean = True) As Boolean
Dim strSql      As String
Dim adoTemp     As ADODB.Recordset
Dim strMensagem As String

    strSql = "SELECT A.intProtocolizacaoVolume Volume, B.strCodigo " & strCONCAT & "'-'" & strCONCAT & "  LTrim(" & gstrCONVERT(CDT_VARCHAR, "B.bitDigito)") & strCONCAT & "'/'" & strCONCAT & "  LTrim(" & gstrCONVERT(CDT_VARCHAR, "B.intExercicio)") & strCONCAT & "' vol.'" & strCONCAT & "  LTrim(" & gstrCONVERT(CDT_VARCHAR, "D.intVolume)") & "  As strProtocolo "
    strSql = strSql & " FROM " & gstrGuiaRemessa & " A " & strREADPAST & ", "
    strSql = strSql & gstrProtocolizacaoProcesso & " B, "
    strSql = strSql & gstrProtocolizacaoVolume & " D " & strREADPAST
    strSql = strSql & " WHERE A.intProtocolizacaoVolume = " & PkidProcesso
    strSql = strSql & " AND D.PKId = A.intProtocolizacaoVolume "
    strSql = strSql & " AND B.PKId = D.intProtocolizacaoProcesso "
    strSql = strSql & " AND A.bytExcluido = 0 "
    strSql = strSql & " ORDER BY A.Pkid DESC"
        
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoTemp) Then
        If Not (adoTemp.BOF And adoTemp.EOF) Then
            ProtocoloEmRemessa = True
            If blnExibeMensagem Then ExibeMensagem " O Processo " & adoTemp!strProtocolo & " está em remessa."
        End If
    End If
    Set gobjBanco = Nothing
                                                        
End Function

Private Function AdicionaCodigo() As String
Dim strCodigo As String
    strCodigo = ""
    
    strCodigo = strCodigo & _
        "Public Function IIf(Expression, TruePart, FalsePart)" & vbCrLf & Chr(9) & _
            "If Expression Then IIf = TruePart Else IIf = FalsePart" & vbCrLf & _
        "End Function" & vbCrLf
    
    strCodigo = strCodigo & _
        "Public Function Format(Expression, sFormat)" & vbCrLf & Chr(9) & _
            "Format = clsRelatorio.Format(Expression, sFormat)" & vbCrLf & _
        "End Function" & vbCrLf
        
    strCodigo = strCodigo & _
        "Public Function CarregaBrasao()" & vbCrLf & Chr(9) & _
            "CarregaBrasao = clsRelatorio.CarregaBrasao" & vbCrLf & _
        "End Function" & vbCrLf
        
    strCodigo = strCodigo & _
        "Public Function CarregaLogotipo()" & vbCrLf & Chr(9) & _
            "CarregaLogotipo = clsRelatorio.CarregaLogotipo" & vbCrLf & _
        "End Function" & vbCrLf
        
    strCodigo = strCodigo & _
        "Public Function CarregaEstado()" & vbCrLf & Chr(9) & _
            "CarregaEstado = clsRelatorio.CarregaEstado" & vbCrLf & _
        "End Function" & vbCrLf
        
    strCodigo = strCodigo & _
        "Public Function CarregaNomeFantasia()" & vbCrLf & Chr(9) & _
            "CarregaNomeFantasia = clsRelatorio.CarregaNomeFantasia" & vbCrLf & _
        "End Function" & vbCrLf
        
    AdicionaCodigo = strCodigo
    
End Function

Public Function gstrQueryTipoCredito(STRTIPO As String, Optional strFiltraDescricao As String) As String
    '-------------------------------------------------------------'
    'valores de bytTipo
    '0 (Orçamentário)
    '1 (Extra-orçamentário)
    '2 (Especial)
    '3 (Suplementar)
    Dim strSql  As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao FROM "
    strSql = strSql & gstrTipoCredito & " "
    strSql = strSql & "WHERE bytTipo IN (" & Trim(STRTIPO) & ")"
    
    If Trim(strFiltraDescricao) <> "" Then
        strSql = strSql & " AND UPPER(strDescricao)  LIKE ('" & UCase(strFiltraDescricao) & "%') "
    End If
    
    strSql = strSql & " ORDER BY  strDescricao "
        
    gstrQueryTipoCredito = strSql
End Function

'Procedimentos para verificar se houve algum dado digitado não salvo
'Argumentos Verifica: True - faz a passagem com os dados no array e verifica se houve alguma mudança
'                    False - Pega os dados do registro de atendimento atual
Public Function VerificaAlteracaoDosDados(Verifica As Boolean, objform As Form, strDados() As String) As Boolean

Dim i   As Integer

With objform
    ReDim Preserve strDados(.Count) As String
    If Not Verifica Then
        For i = 0 To .Controls.Count - 1
            If TypeOf .Controls(i) Is TextBox Or TypeOf .Controls(i) Is DataCombo Then
                strDados(i) = .Controls(i).Text
            End If
        Next i
    Else
        For i = 0 To .Controls.Count - 1
            If TypeOf .Controls(i) Is TextBox Or TypeOf .Controls(i) Is DataCombo Then
                If strDados(i) <> .Controls(i).Text Then
                    If MsgBox("Deseja Abandonar?", vbQuestion + vbYesNo, "Dados não atualizados") = vbNo Then
                        VerificaAlteracaoDosDados = True
                        Exit Function
                    Else
                        VerificaAlteracaoDosDados = False
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
End With

End Function

'Public Function SaldoDotacaoAtual(intDotacao As Double, intMes As Integer, intExercicio As Integer, Optional DBLVALOR As Double) As Double
'
'    'Esta função está sendo utilizada pelos módulos Orçamentário e Compras
'    'Movimentada do ModOrcamentario para ModGeral em 08/12/03 por Alfred
'
'   Dim strSql          As String
'   Dim intCont         As Integer
'   Dim dblSaldoDotacao      As Double
'   Dim dblSaldo(1 To 12, 0 To 6) As Double
'   Dim dblSaldoIni     As Double
'   Dim dblSuplementado As Double
'   Dim dblAnulado      As Double
'   Dim dblEmpenhado    As Double
'   Dim dblReservado    As Double
'   Dim dblBloqueado    As Double
'   Dim adoResultado    As New ADODB.Recordset
'
'
'   'Vamos buscar o valor do Saldo Inicial
'
'   strSql = "SELECT dblValor FROM " & gstrProgramaDeTrabalho
'   strSql = strSql & " WHERE PKID = " & intDotacao & " AND intExercicio = " & intExercicio
'
'   Set gobjBanco = New clsBanco
'
'   If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'    If Not adoResultado.EOF Then dblSaldoIni = IIf(IsNull(adoResultado!DBLVALOR), gstrConvVrDoSql("0", 2, 5), adoResultado!DBLVALOR)
'   End If
'   adoResultado.Close
'
'   SaldoDotacaoAtual = dblSaldoIni
'   dblSaldoDotacao = dblSaldoIni
'
'   For intCont = 1 To 12
'
'       strSql = "SELECT " & gstrISNULL("SUM(DSR.dblValor)", "0") & " AS dblValor FROM " & gstrSuplementacaoReducao & " SR, " & gstrDotacaoSuplementadaReduzida & " DSR "
'       strSql = strSql & " WHERE SR.PKID = DSR.intSuplementacaoReducao AND DSR.intProgramaTrabalho = " & intDotacao
'       strSql = strSql & " AND DSR.bytOperacao = 2 AND " & gstrDATEPART(strYEAR, "SR.dtmDataDecreto") & " = " & gintExercicio
'       strSql = strSql & " AND " & gstrDATEPART(strMONTH, "SR.dtmDataDecreto") & " = " & intCont
'
'       If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'          dblSuplementado = adoResultado!DBLVALOR
'       End If
'
'       adoResultado.Close
'
'       strSql = "SELECT " & gstrISNULL("SUM(SRD.dblValor)", "0") & " AS dblValor FROM " & gstrSuplementacaoReducao & " SR, " & gstrSuplementacaoReducaoDespesa & " SRD "
'       strSql = strSql & " WHERE SR.PKID = SRD.intSuplementacaoReducao AND SRD.intProgramaTrabalho = " & intDotacao
'       strSql = strSql & " AND " & gstrDATEPART(strYEAR, "SR.dtmDataDecreto") & " = " & gintExercicio
'       strSql = strSql & " AND " & gstrDATEPART(strMONTH, "SR.dtmDataDecreto") & " = " & intCont
'
'       If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'          dblSuplementado = dblSuplementado + adoResultado!DBLVALOR
'       End If
'
'       adoResultado.Close
'
'       strSql = "SELECT " & gstrISNULL("SUM(DSR.dblValor)", "0") & " AS dblValor FROM " & gstrSuplementacaoReducao & " SR, " & gstrDotacaoSuplementadaReduzida & " DSR "
'       strSql = strSql & " WHERE SR.PKID = DSR.intSuplementacaoReducao AND DSR.intProgramaTrabalho = " & intDotacao
'       strSql = strSql & " AND DSR.bytOperacao = 1 AND " & gstrDATEPART(strYEAR, "SR.dtmDataDecreto") & " = " & gintExercicio
'       strSql = strSql & " AND " & gstrDATEPART(strMONTH, "SR.dtmDataDecreto") & " = " & intCont
'
'       If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'          dblAnulado = adoResultado!DBLVALOR
'       End If
'
'       adoResultado.Close
'
'       strSql = "SELECT " & gstrISNULL("SUM(dblValor)", "0") & " AS dblValor FROM " & gstrEmpenho
'       strSql = strSql & " WHERE intReservaDotacao IS NULL AND intProgramaTrabalho = " & intDotacao & " AND "
'       strSql = strSql & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio & " AND "
'       strSql = strSql & gstrDATEPART(strMONTH, "dtmData") & " = " & intCont
'
'       If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'          dblEmpenhado = adoResultado!DBLVALOR
'       End If
'
'       adoResultado.Close
'
'       strSql = "SELECT " & gstrISNULL("SUM(SEP.dblValor)", "0") & " AS dblValor FROM " & gstrSubempenho & " SEP, "
'       strSql = strSql & gstrEmpenho & " EP "
'       strSql = strSql & " WHERE EP.intReservaDotacao IS NULL AND EP.PKID = SEP.intEmpenho AND "
'       strSql = strSql & " EP.intProgramaTrabalho = " & intDotacao & " AND "
'       strSql = strSql & " SEP.intNumero = 0 AND bytSituacao = 4 AND "
'       strSql = strSql & gstrDATEPART(strYEAR, "SEP.dtmData") & " = " & gintExercicio & " AND "
'       strSql = strSql & gstrDATEPART(strMONTH, "SEP.dtmData") & " = " & intCont
'
'       If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'          dblEmpenhado = dblEmpenhado - adoResultado!DBLVALOR
'       End If
'
'       adoResultado.Close
'
'
'       strSql = "SELECT " & gstrISNULL("SUM(dblValor)", "0") & " AS dblValor FROM " & gstrReservaDotacao
'       strSql = strSql & " WHERE intProgramaTrabalho = " & intDotacao & " AND "
'       strSql = strSql & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio & " AND "
'       strSql = strSql & gstrDATEPART(strMONTH, "dtmData") & " = " & intCont
'
'       If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'          dblReservado = adoResultado!DBLVALOR
'       End If
'
'       adoResultado.Close
'
'       strSql = "SELECT " & gstrISNULL("SUM(RDL.dblValor)", "0") & " AS dblValor FROM " & gstrReservaDotacaoLiberada & " RDL,"
'       strSql = strSql & gstrReservaDotacao & " RD WHERE RD.intProgramaTrabalho = " & intDotacao
'       strSql = strSql & " AND RD.PKID = RDL.intReservaDotacao AND "
'       strSql = strSql & gstrDATEPART(strYEAR, "RDL.dtmData") & " = " & gintExercicio & " AND "
'       strSql = strSql & gstrDATEPART(strMONTH, "RDL.dtmData") & " = " & intCont & " AND "
'       strSql = strSql & " (RDL.intFlag = 0 OR RDL.intFlag = 1 AND RDL.dblValor < 0)"
'
'       If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'          dblReservado = dblReservado - adoResultado!DBLVALOR
'       End If
'
'       adoResultado.Close
'
'
'       strSql = "SELECT " & gstrISNULL("SUM(dblValor)", "0") & " AS dblValor FROM " & gstrContencaoCredito
'       strSql = strSql & " WHERE intPrograma = " & intDotacao & " AND "
'       strSql = strSql & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio & " AND "
'       strSql = strSql & gstrDATEPART(strMONTH, "dtmData") & " = " & intCont
'
'       If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'          dblBloqueado = adoResultado!DBLVALOR
'       End If
'
'       adoResultado.Close
'
'       strSql = "SELECT " & gstrISNULL("SUM(CCD.dblValor)", "0") & " AS dblValor FROM " & gstrContencaoCreditoDesbloqueado & " CCD,"
'       strSql = strSql & gstrContencaoCredito & " CC "
'       strSql = strSql & " WHERE CC.PKID = CCD.intContencao AND "
'       strSql = strSql & " CC.intPrograma = " & intDotacao & " AND "
'       strSql = strSql & gstrDATEPART(strYEAR, "CCD.dtmDesbloqueio") & " = " & gintExercicio & " AND "
'       strSql = strSql & gstrDATEPART(strMONTH, "CCD.dtmDesbloqueio") & " = " & intCont
'
'       If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'          dblBloqueado = dblBloqueado - adoResultado!DBLVALOR
'       End If
'
'       adoResultado.Close
'
'       dblSaldo(intCont, 0) = dblSaldoIni
'       dblSaldo(intCont, 1) = dblSuplementado
'       dblSaldo(intCont, 2) = dblAnulado
'       dblSaldo(intCont, 3) = dblEmpenhado
'       dblSaldo(intCont, 4) = dblReservado
'       dblSaldo(intCont, 5) = dblBloqueado
'       dblSaldo(intCont, 6) = 0
'
'   Next
'
'   For intCont = 1 To intMes
'       dblSaldoDotacao = dblSaldoDotacao + ((dblSaldo(intCont, 1) - dblSaldo(intCont, 2)) - (dblSaldo(intCont, 3) + dblSaldo(intCont, 4) + dblSaldo(intCont, 5)))
'   Next
'
'   SaldoDotacaoAtual = dblSaldoDotacao
'
'   If DBLVALOR > 0 Then
'      If DBLVALOR > SaldoDotacaoAtual Then
'         ExibeMensagem "Saldo insulficiente para este movimento."
'         SaldoDotacaoAtual = Empty
'         Exit Function
'      Else
'         For intMes = intMes + 1 To 12
'
'            SaldoDotacaoAtual = SaldoDotacaoAtual + ((dblSaldo(intMes, 1) - dblSaldo(intMes, 2)) - (dblSaldo(intMes, 3) + dblSaldo(intMes, 4) + dblSaldo(intMes, 5)))
'
'            If DBLVALOR > SaldoDotacaoAtual Then
'               ExibeMensagem "Saldo de dotação insuficiente nos meses posteriores."
'               SaldoDotacaoAtual = Empty
'               Exit Function
'            End If
'
'         Next
'      End If
'   End If
'
'   SaldoDotacaoAtual = dblSaldoDotacao
'
'End Function

Public Function strPosicaoTamanho(bytSetorOuQuadra As Byte) As String
'Nino - Função para retornar a posicao e o tamanho que deve se obter os valores de setor e quadra da máscara
Dim strSql       As String
Dim adoResultado As ADODB.Recordset

If bytSetorOuQuadra = 1 Then
    strPosicaoTamanho = 1
ElseIf bytSetorOuQuadra = 2 Then
    strSql = "SELECT SUM(CI.intTamanho + " & strLen & "(CI.strSeparador)) AS Posicao"
    strSql = strSql & " FROM "
    strSql = strSql & gstrCampoDeInscricao & " CI "
    strSql = strSql & " WHERE "
    strSql = strSql & " CI.bytSetorQuadra = " & bytSetorOuQuadra '1 Setor - 2 Quadra
    strSql = strSql & " AND CI.inttipodeinscricao = " & TYP_IMOBILIARIA
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            If Not IsNull(adoResultado!Posicao) Then
                If Trim(Val(gstrENulo(adoResultado!Posicao))) = 1 Then
                    strPosicaoTamanho = 0
                Else
                    strPosicaoTamanho = Trim(Val(gstrENulo(adoResultado!Posicao)))
                End If
            Else
                strPosicaoTamanho = 0
            End If
        Else
            strPosicaoTamanho = 0
        End If
    End If
End If

strSql = "SELECT intTamanho Tamanho"
strSql = strSql & " FROM "
strSql = strSql & gstrCampoDeInscricao
strSql = strSql & " WHERE bytSetorQuadra = " & bytSetorOuQuadra & " and inttipodeinscricao = " & TYP_IMOBILIARIA

If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
    If Not adoResultado.EOF Then
        strPosicaoTamanho = strPosicaoTamanho & ";" & adoResultado!Tamanho
    End If
End If

End Function

Public Function VerificaMaterial(dbcCodigo As DataCombo, dbcDescricao As DataCombo) As Boolean
Dim adoDesc          As ADODB.Recordset
Dim strSql           As String
Dim strGrupo         As String
Dim STRTIPO          As String
Dim strFamilia       As String
Dim strItem          As String
Dim strCodMaterial   As String
    
    VerificaMaterial = False
    
    If dbcCodigo.Text = "" Then
        Exit Function
    Else
        
        If dbcCodigo.MatchedWithList And dbcDescricao.MatchedWithList Then
            Exit Function
        End If
        
        If bitCodigoCompleto = vbUnchecked Then
            
            strSql = "SELECT PKid, strDescricao"
            strSql = strSql & " FROM " & gstrCatalogoMaterialServico
            strSql = strSql & " WHERE intCodigo = " & dbcCodigo.Text
        
        Else
        
            strCodMaterial = Replace(dbcCodigo.Text, ".", "")
            
            strGrupo = Val(Left$(strCodMaterial, intTamanhoGrupo))
            STRTIPO = Val(Mid$(strCodMaterial, intTamanhoGrupo + 1, intTamanhoTipo))
            
            If (Len(strCodMaterial) = (intTamanhoGrupo + intTamanhoTipo + intTamanhoFamilia + intTamanhoItem)) And (intTamanhoFamilia > 0) Then
                strFamilia = Val(Mid$(strCodMaterial, intTamanhoGrupo + intTamanhoTipo + 1, intTamanhoFamilia))
                strItem = Val(Right$(strCodMaterial, intTamanhoItem))
            Else
                strItem = Val(Right$(strCodMaterial, intTamanhoItem))
            End If
            
            strSql = "SELECT CM.PKid, CM.strDescricao"
            strSql = strSql & " FROM " & gstrCatalogoMaterialServico & " CM, "
            strSql = strSql & gstrGrupoMaterialServico & " GM, "
            strSql = strSql & gstrTipoMaterialServico & " TM, "
            strSql = strSql & gstrFamiliaMaterial & " FM "
            
            strSql = strSql & "WHERE CM.intGrupoMaterialServico = GM.PKid AND "
            strSql = strSql & "CM.intTipoMaterialServico = TM.PKid AND "
            strSql = strSql & "CM.intFamiliaMaterial " & strOUTJSQLServer & "= FM.PKid " & strOUTJOracle
            
            strSql = strSql & " AND GM.intCodigo = " & strGrupo & " AND"
            strSql = strSql & " TM.intCodigo = " & STRTIPO & " AND "
            
            If (Len(strCodMaterial) = (intTamanhoGrupo + intTamanhoTipo + intTamanhoFamilia + intTamanhoItem)) And (intTamanhoFamilia > 0) Then
                strSql = strSql & "FM.strCodigo = " & strFamilia & " AND "
                strSql = strSql & "CM.intCodigo = " & strItem
            Else
                strSql = strSql & "CM.intCodigo = " & strItem
            End If
            
        End If
        
        Set gobjBanco = New clsBanco
        
        If gobjBanco.CriaADO(strSql, 30, adoDesc) Then
            With adoDesc
                If Not .EOF And Not .BOF Then
                    PreencherListaDeOpcoes dbcCodigo, !Pkid
                    PreencherListaDeOpcoes dbcDescricao, !Pkid
                
                    If bitCodigoCompleto = vbChecked Then
                        dbcCodigo.Text = strGrupo & "." & STRTIPO & "." & IIf(Trim$(strFamilia) <> Space$(0), strFamilia & "." & strItem, strItem)
                    End If
                
                Else
                    dbcDescricao.BoundText = ""
                    dbcDescricao.Text = ""
                    
                    dbcCodigo.Text = ""
                    dbcDescricao.Text = ""
                    
                    VerificaMaterial = True
                    
                End If
            End With
        End If
            
    End If
        
    DoEvents
    
End Function

Public Sub VerificaParametroMateriais()
Dim strSql       As String
Dim adoParametro As ADODB.Recordset

    strSql = "SELECT * FROM " & gstrParametroMateriais
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 30, adoParametro) Then
        With adoParametro
            If Not .EOF And Not .BOF Then
                intTamanhoGrupo = gstrENulo(!intTamanhoGrupo)
                intTamanhoTipo = gstrENulo(!intTamanhoTipo)
                intTamanhoFamilia = gstrENulo(!intTamanhoFamilia)
                intTamanhoItem = gstrENulo(!intTamanhoItem)
                bitCodigoCompleto = IIf(gstrENulo(!bitCodigoCompleto), vbChecked, vbUnchecked)
            End If
        End With
    End If
    
End Sub

Public Sub AtualizaParametroPedidoDigital()
Dim strSql       As String
Dim adoParametro As ADODB.Recordset

    strSql = "SELECT * FROM " & gstrParametrosEspecificos
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 30, adoParametro) Then
        With adoParametro
            If Not .EOF And Not .BOF Then
                gstrTitulolPedidoCotacao = !strTitulolPedidoCotacao & Space$(0)
                gstrFornecedor = !strFornecedor & Space$(0)
                gstrIdentificacao = !strIdentificacao & Space$(0)
                gstrEndereco = !STRENDERECO & Space$(0)
                gstrTelefone = !strTelefone & Space$(0)
                gstrContato = !strContato & Space$(0)
                gstrObservacao = !strObservacao & Space$(0)
                gintLinhaInicial = IIf(IsNull(!intLinhaInicial), 0, !intLinhaInicial)
                gstrSenha = gstrStringCripitografada(!strSenha & Space$(0), , True)
                gintColNumeroDoItem = IIf(IsNull(!intColNumeroDoItem), 0, !intColNumeroDoItem)
                gintColDescricaoDoItem = IIf(IsNull(!intColDescricaoDoItem), 0, !intColDescricaoDoItem)
                gintColComplementoDoItem = IIf(IsNull(!intColComplementoDoItem), 0, !intColComplementoDoItem)
                gintColMarca = IIf(IsNull(!intColMarca), 0, !intColMarca)
                gintColUnidadeDeMedida = IIf(IsNull(!intColUnidadeDeMedida), 0, !intColUnidadeDeMedida)
                gintColQuantidade = IIf(IsNull(!intColQuantidade), 0, !intColQuantidade)
                gintColValorUnitario = IIf(IsNull(!intColValorUnitario), 0, !intColValorUnitario)
                gintColValorTotal = IIf(IsNull(!intColValorTotal), 0, !intColValorTotal)
            End If
        End With
        adoParametro.Close
    End If
    
    Set adoParametro = Nothing
    
End Sub

Public Sub CarregaParametrosEspecificos()
Dim strSql       As String
Dim adoParametro As ADODB.Recordset

    strSql = "SELECT bytReserva FROM " & gstrParametrosEspecificos
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 30, adoParametro) Then
        With adoParametro
            If Not .EOF And Not .BOF Then
                bytParametroReserva = gstrENulo(!bytReserva)
            End If
        End With
    End If

End Sub

Public Sub ImprimeRelatorioGrid(Grid As TDBGrid)
    
    Set GridDeImpressao = Grid
    
    GridDeImpressao.PrintInfo.PreviewInitZoom = 100
    GridDeImpressao.PrintInfo.PreviewMaximize = True

    MDIMenu.PopupMenu MDIMenu.mnuImprimir
    
End Sub

Public Function VerificaFechamentos(lngAlmoxarifado As Long, dtmDataComparacao As Date, Optional strDescricaoAlmoxarifado As String) As Boolean
Dim adoRec As New ADODB.Recordset

    VerificaFechamentos = False
    
    Set gobjBanco = New clsBanco
    
    'Vamos verificar o fechamento mensal
    If gobjBanco.CriaADO("SELECT MAX(dtmData) AS dtmData FROM " & gstrFechamentoMateriais & " WHERE intAlmoxarifado = " & lngAlmoxarifado, 100, adoRec) Then
        If Not IsNull(adoRec!DTMDATA) Then
            If CDate(adoRec!DTMDATA) >= CDate(dtmDataComparacao) Then
                ExibeMensagem "A data de movimentação tem que ser superior à data de fechamento do almoxarifado." & vbCrLf & Trim(strDescricaoAlmoxarifado) & " - Fechamento: " & gstrDataFormatada(adoRec!DTMDATA)
                Exit Function
            End If
        End If
    End If
    
    'Vamos verificar o fechamento de inventario
    If gobjBanco.CriaADO("SELECT MAX(dtmDatadeBloqueio) AS dtmData FROM " & gstrInventarioMaterial & " WHERE intAlmoxarifado = " & lngAlmoxarifado & " AND bytAbertoFechado = 0", 100, adoRec) Then
        If Not IsNull(adoRec!DTMDATA) Then
            If CDate(adoRec!DTMDATA) >= CDate(dtmDataComparacao) Then
                ExibeMensagem "A data de movimentação tem que ser superior à data de bloqueio de inventário." & vbCrLf & Trim(strDescricaoAlmoxarifado) & " - Bloqueio: " & gstrDataFormatada(adoRec!DTMDATA)
                Exit Function
            End If
        End If
    End If
    
    VerificaFechamentos = True
    
End Function

Public Function GravaReservaDotacao(blnPorAutorizacao As Boolean, lngCodigo As Long, Optional intExercicio As Integer, _
                                Optional strObjetoAutorizacao As String, Optional strComprasLicitacao As String)
                                
Dim strSql      As String
Dim mstrCodigo  As String
Dim adoRec      As ADODB.Recordset
Dim lngContabil As Long
    
    GravaReservaDotacao = False
    
    strSql = ""

    strSql = strSql & "SELECT "
    strSql = strSql & "A.IntProgramaDeTrabalho Contabilidade, B.strCodigo, A.intCodigo, A.intExercicio, "
    strSql = strSql & "(SELECT SUM(B.DblQuantidade * B.dblValorEstimado) FROM " & gstrRequisicaoCompras & " B WHERE B.intCodigo = A.intCodigo) TOTAL "
    strSql = strSql & "FROM " & gstrRequisicaoCompras & " A, " & gstrProgramaDeTrabalho & " B "
    strSql = strSql & "WHERE (NOT a.IntProgramaDeTrabalho IS NULL AND a.intPlanoConta IS NULL) AND intReserva IS NULL "
    
    If blnPorAutorizacao Then
        strSql = strSql & " AND A.intAutorizacaoDeCompra = " & lngCodigo
    Else
        strSql = strSql & " AND A.intCodigo = " & lngCodigo & " AND A.intExercicio = " & intExercicio
    End If
        
    strSql = strSql & " AND B.Pkid = A.IntProgramaDeTrabalho "
    strSql = strSql & " GROUP BY A.intCodigo, A.IntProgramaDeTrabalho, B.strCodigo, A.intExercicio "
    strSql = strSql & "ORDER BY Contabilidade "
        
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
        If Not (adoRec.EOF And adoRec.BOF) Then
                
            If MsgBox("Deseja cadastrar a reserva ?", vbYesNo + vbQuestion) = vbNo Then
                Exit Function
            End If
                
            lngContabil = 0
            
            gobjBanco.ExecutaBeginTrans
            
            Do While Not adoRec.EOF
               
                If VerificaSaldo(adoRec!intCodigo, adoRec!intExercicio) Then
    
                    If adoRec!Contabilidade <> lngContabil Then
                    
                        strSql = ""
                        strSql = strSql & "INSERT INTO " & gstrReservaDotacao & " "
                
                        strSql = strSql & "(intNumero, dtmData, dblValor, intProgramaTrabalho, "
                        strSql = strSql & "strHistorico, strSolicitante, strSolicitacao, "
                        strSql = strSql & "dtmDtAtualizacao, lngCodUsr) "
              
                        strSql = strSql & "SELECT MAX(intNumero) + 1, "
                        strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                        strSql = strSql & gstrConvVrParaSql(adoRec!TOTAL) & ", "
                       
                        strSql = strSql & adoRec!Contabilidade & ", "
                        
                        strSql = strSql & "'" & Trim(strObjetoAutorizacao) & "', "
                        strSql = strSql & "'', "
                        strSql = strSql & "'" & Trim(strComprasLicitacao) & "', "
                        strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                        strSql = strSql & glngCodUsr & " FROM " & gstrReservaDotacao
                    
                        gobjBanco.Execute strSql
                                         
                        lngContabil = adoRec!Contabilidade
                        
                        strSql = ""
                        strSql = strSql & "SELECT " & gstrTOPnSQLServer(1) & " PKid "
                        strSql = strSql & "FROM " & gstrReservaDotacao
                        strSql = strSql & " WHERE lngCodUsr = " & glngCodUsr
                        strSql = strSql & " ORDER BY PKid DESC"
                        strSql = gstrTOPnOracle(strSql, 1)
                        
                        strSql = "UPDATE " & gstrRequisicaoCompras & " SET intReserva = (" & strSql & ")"
                        strSql = strSql & "WHERE "
                        If blnPorAutorizacao Then
                            strSql = strSql & "intProgramaDeTrabalho = " & adoRec!Contabilidade & " AND intAutorizacaoDeCompra = " & lngCodigo
                        Else
                            strSql = strSql & "intCodigo = " & lngCodigo & " AND intExercicio = " & intExercicio
                        End If
                    
                        gobjBanco.Execute strSql
                        
                        strSql = ""
                        strSql = "UPDATE " & gstrRequisicaoCompras
                        strSql = strSql & " SET intRequisicaoComprasSituacoes = (SELECT pkid FROM " & gstrRequisicaoComprasSituacoes & " WHERE bitTipo = 3)"
                        
                        gobjBanco.Execute strSql
                        
                    End If
                    
                    adoRec.MoveNext
                        
                Else
                    ExibeMensagem "Saldo insulficiente para o Programa de Trabalho " & adoRec!strCodigo & "."
                    gobjBanco.ExecutaRollbackTrans
                    GravaReservaDotacao = False
                    Exit Function
                End If
                
            Loop
                        
            gobjBanco.ExecutaCommitTrans
            
            GravaReservaDotacao = True
            
        Else
            ExibeMensagem "É necessário relacionar a Solicitação de Compra com algum Programa de Trabalho e não possuir Reserva de Dotação."
            Exit Function
        End If
        
    End If
        
    Set gobjBanco = Nothing
    
End Function

Public Sub RemoveReservaDotacao(lngCodigo As Long, intExercicio As Integer)

    Dim strSql      As String
    Dim adoTemp     As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & " SELECT " & gstrTOPnSQLServer(1) & " intReserva "
    strSql = strSql & " FROM " & gstrRequisicaoCompras
    strSql = strSql & " WHERE intCodigo = " & lngCodigo & " AND intExercicio = " & intExercicio
    strSql = gstrTOPnOracle(strSql, 1)
    
    strSql = " SELECT * FROM " & gstrReservaDotacao & " WHERE PKid = (" & strSql & ")"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoTemp) Then
    
        If Not (adoTemp.EOF And adoTemp.BOF) Then
            
            strSql = ""
            strSql = strSql & " INSERT INTO " & gstrReservaDotacaoLiberada & " ("
            strSql = strSql & " intReservaDotacao, "
            strSql = strSql & " intNumero, "
            strSql = strSql & " dtmData, "
            strSql = strSql & " dblValor, "
            strSql = strSql & " strHistorico, "
            strSql = strSql & " dtmDtAtualizacao, "
            strSql = strSql & " lngCodUsr, "
            strSql = strSql & " intFlag) "
            strSql = strSql & " SELECT " & adoTemp!Pkid & ", "
            strSql = strSql & gstrCASEWHEN("COUNT(PKid)", "0, 1", "MAX(intNumero) + 1") & ", "
            strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strSql = strSql & gstrConvVrParaSql(adoTemp!dblValor) & ", "
            strSql = strSql & "NULL, "
            strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strSql = strSql & glngCodUsr & ", "
            strSql = strSql & "0 "
            strSql = strSql & " FROM " & gstrReservaDotacaoLiberada & " WHERE intReservaDotacao = " & adoTemp!Pkid
            
            If gobjBanco.Execute(strSql) Then
                
                strSql = ""
                strSql = strSql & " UPDATE " & gstrRequisicaoCompras & " SET "
                strSql = strSql & " intReserva = NULL "
                strSql = strSql & " WHERE intCodigo = " & lngCodigo & " AND intExercicio = " & intExercicio
                
                gobjBanco.Execute strSql
                
            End If
            
        End If
        
    End If
    
    Set gobjBanco = Nothing

End Sub

Private Function VerificaSaldo(lngCodigo As Long, intExercicio As Integer) As Boolean
Dim adoRec  As ADODB.Recordset
    
    VerificaSaldo = True
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO("SELECT RC.intProgramadeTrabalho, RC.dtmDataRequisicao, SUM(RC.dblQuantidade * RC.dblValorEstimado) AS Total FROM " & gstrRequisicaoCompras & " RC WHERE RC.intCodigo = " & lngCodigo & " AND RC.intExercicio = " & intExercicio & " GROUP BY RC.intProgramadeTrabalho , RC.dtmDataRequisicao", 0, adoRec) Then
        
        Set gobjBanco = Nothing

        With adoRec
        
            If Not IsNull(!intProgramadeTrabalho) Then
                If !TOTAL > SaldoDotacaoAtual(!intProgramadeTrabalho, Month(!dtmDataRequisicao), Year(!dtmDataRequisicao)) Then
                    VerificaSaldo = False
                End If
            End If
        
        End With

    End If

End Function

Public Function SaldoDotacaoAtual(lngDotacao As Long, intMes As Integer, intExercicio As Integer, Optional dblValor As Double) As Double
    'Esta função está sendo utilizada pelos módulos Orçamentário e Compras
    'Movimentada do ModOrcamentario para ModGeral em 08/12/03 por Alfred
    
Dim strSql                       As String
Dim intCont                      As Integer
Dim dblSaldoDotacao              As Double
Dim dblSaldo(1 To 12, 0 To 6)    As Double
Dim dblSaldoIni                  As Double
Dim dblSuplementado              As Double
Dim dblAnulado                   As Double
Dim dblEmpenhado                 As Double
Dim dblReservado                 As Double
Dim dblBloqueado                 As Double
Dim adoResultado                 As ADODB.Recordset
   'Suplemento - S
   'Anulado    - A
   'Empenhado  - E
   'Reservado  - R
   'Bloqueado  - B
   
   
   'Vamos buscar o valor do Saldo Inicial

    strSql = "SELECT dblValor FROM " & gstrProgramaDeTrabalho
    strSql = strSql & " WHERE PKID = " & lngDotacao & " AND intExercicio = " & intExercicio
   
    Set gobjBanco = New clsBanco
   
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then dblSaldoIni = IIf(IsNull(adoResultado!dblValor), gstrConvVrDoSql("0", 2, 5), adoResultado!dblValor)
    End If
   
   
    SaldoDotacaoAtual = dblSaldoIni
    dblSaldoDotacao = dblSaldoIni
    
    strSql = ""
    strSql = strSql & " SELECT saldoIni, SUM(Empenhado) dblEmpenhado , "
    strSql = strSql & " SUM(Suplementado) dblSuplementado,"
    strSql = strSql & " SUM(Anulado) dblAnulado,SUM(Bloqueado) dblBloqueado,"
    strSql = strSql & " SUM(Reservado) dblReservado, mes "
    strSql = strSql & " FROM " & gstrContaValoresAcumulados & " "
    strSql = strSql & " WHERE intProgramadeTrabalho = " & lngDotacao
    strSql = strSql & " AND intExercicio = " & gintExercicio
    strSql = strSql & " GROUP BY saldoIni, mes"

           
        For intCont = 1 To 12
        dblSaldo(intCont, 0) = 0
        dblSaldo(intCont, 1) = 0
        dblSaldo(intCont, 2) = 0
        dblSaldo(intCont, 3) = 0
        dblSaldo(intCont, 4) = 0
        dblSaldo(intCont, 5) = 0
        dblSaldo(intCont, 6) = 0
    Next
       
    Set gobjBanco = New clsBanco
       
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            adoResultado.MoveFirst
            For intCont = 1 To 12
                With adoResultado
                    dblSaldoDotacao = !SaldoIni
                    
                    dblSaldo(intCont, 1) = !dblSuplementado - !dblAnulado
                    dblSuplementado = 0
                    
                    dblSaldo(intCont, 3) = !dblEmpenhado
                    dblEmpenhado = 0
                    
                    dblSaldo(intCont, 4) = !dblReservado
                    dblReservado = 0
                    
                    dblSaldo(intCont, 5) = !dblBloqueado
                    dblBloqueado = 0
                    .MoveNext
                End With
            Next
        End If
    End If
       
    For intCont = 1 To intMes
        dblSaldoDotacao = dblSaldoDotacao + ((dblSaldo(intCont, 1) - dblSaldo(intCont, 2)) - (dblSaldo(intCont, 3) + dblSaldo(intCont, 4) + dblSaldo(intCont, 5)))
    Next
   
    SaldoDotacaoAtual = dblSaldoDotacao
   
    If dblValor > 0 Then
        If dblValor > SaldoDotacaoAtual Then
            ExibeMensagem "Saldo insulficiente para este movimento."
            SaldoDotacaoAtual = Empty
            Exit Function
        Else
            For intMes = intMes + 1 To 12
                SaldoDotacaoAtual = SaldoDotacaoAtual + ((dblSaldo(intMes, 1) - dblSaldo(intMes, 2)) - (dblSaldo(intMes, 3) + dblSaldo(intMes, 4) + dblSaldo(intMes, 5)))
                If dblValor > SaldoDotacaoAtual Then
                    ExibeMensagem "Saldo de dotação insuficiente nos meses posteriores."
                    SaldoDotacaoAtual = Empty
                    Exit Function
                End If
             Next
        End If
   End If
   
   SaldoDotacaoAtual = dblSaldoDotacao
   
End Function

Public Sub AtribuiSituacaoInicial()
Dim strSql As String
Dim adoRec As ADODB.Recordset

    strSql = "SELECT * FROM " & gstrRequisicaoComprasSituacoes
    
    Set gobjBanco = New clsBanco
    gobjBanco.CriaADO strSql, 5, adoRec
    With adoRec
    
        Do While Not .EOF
        
            Select Case !strDescricao
                
                Case Is = SIGLA_EMCADASTRAMENTO
                    EMCAD = !Pkid
            
            End Select
            
            .MoveNext
        
        Loop
    
    End With
            
End Sub

Public Function blnConverteUnidade(lngCatalogoMaterial As Long, lngUnidadeMedida As Long) As Double
Dim strSql       As String
Dim adoResultado As ADODB.Recordset

    strSql = "SELECT " & gstrISNULL("intQuantidade", "1") & " FROM " & gstrCatalogoMaterialServicoUnid
    strSql = strSql & " WHERE intUnidadeMedida ='" & lngUnidadeMedida & "'"
    strSql = strSql & " AND intCatalogoMaterialServico ='" & lngCatalogoMaterial & "'"
    
    Set gobjBanco = New clsBanco

    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF Then
            blnConverteUnidade = 1
        Else
            blnConverteUnidade = adoResultado(0).Value
        End If
    End If
    
    Set gobjBanco = Nothing

End Function

Public Function VerificaEmpenhoProcesso(strCodigo As String, bitDigito As Integer, intExercicio As Integer) As Boolean
   Dim strSql As String
   Dim adoResultado As New ADODB.Recordset
   
   strSql = "SELECT PP.PKID FROM " & gstrProtocolizacaoProcesso & " PP "
   strSql = strSql & "WHERE Ltrim(Rtrim(PP.strCodigo)) = '" & Trim(strCodigo) & "' AND "
   strSql = strSql & IIf(Len(Trim(bitDigito)) > 0, "Ltrim(Rtrim(PP.bitDigito)) =  " & Trim(bitDigito), "bitDigito IS NULL") & " AND "
   strSql = strSql & IIf(Len(Trim(intExercicio)) > 0, "Ltrim(Rtrim(PP.intExercicio)) =  " & Trim(intExercicio), "intExercicio IS NULL")
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
      If Not adoResultado.EOF Then
         VerificaEmpenhoProcesso = True
      End If
   End If
   
End Function

Public Sub ImprimeRelatorioDoGrid(objRelatorio As Object, _
                            strQuery As String, _
                            strConnectionString As String, _
                            adoRelatorio As ADODB.Recordset, _
                            tdbGridRelatorio As TDBGrid, _
                   Optional strTitulo As String, _
                   Optional lngIntervaloDeTempo As Long)
    
    
    Dim strRelatorio            As String
    Dim objActiveReports        As ActiveReport
    Dim blnExisteArquivo        As Boolean
    On Error GoTo ErroImprimeRelatorio
    
    If Trim(strQuery) = "" Then
        Exit Sub
    End If
    If lngIntervaloDeTempo = 0 Then
        lngIntervaloDeTempo = 30
    End If
    Screen.MousePointer = vbHourglass
    
    strRelatorio = objRelatorio.Name & ".rpx"
    
    gblnRestartRelatorio = False
    Set gobjBanco = New clsBanco
    
    On Error GoTo NaoExiste
    If Dir(gstrDirDocumentos & "Documentos\Relatorios\" & strRelatorio, vbArchive) <> "" Then
        blnExisteArquivo = True
    Else
NaoExiste:
        blnExisteArquivo = False
    End If
    
    On Error GoTo ErroImprimeRelatorio

    If blnExisteArquivo Then
        Set objActiveReports = New ActiveReport
        
        objActiveReports.LoadLayout gstrDirDocumentos & "Documentos\Relatorios\" & strRelatorio
        objActiveReports.adoDataControl.ConnectionString = strConnectionString
        objActiveReports.adoDataControl.Source = strQuery
        frmVisualizarRelatorio.ARViewer.ReportSource = objActiveReports
        If Trim(strTitulo) <> "" Then
            frmVisualizarRelatorio.Caption = strTitulo
        End If
        
        objActiveReports.ResetScripts
        
        objActiveReports.AddCode AdicionaCodigo
        
        objActiveReports.AddNamedItem "clsRelatorio", New clsRelatorio
        
        frmVisualizarRelatorio.WindowState = vbMaximized
        frmVisualizarRelatorio.Show
    Else
        objRelatorio.adoDataControl.Provider = ""
        objRelatorio.adoDataControl.ConnectionString = strConnectionString
        objRelatorio.adoDataControl.Source = strQuery
        PadronizaRelatorio objRelatorio, adoRelatorio, tdbGridRelatorio
        Set objRelatorio.adoDataControl.Recordset = adoRelatorio
        If Trim(strTitulo) <> "" Then
            objRelatorio.Caption = strTitulo
        End If
        objRelatorio.WindowState = vbMaximized
        
        objRelatorio.Show
        
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
    
ErroImprimeRelatorio:
    ExibeDetalheErro "Erro na geração do relatório."
    Resume FimImprimeRelatorio
    
FimImprimeRelatorio:
    Screen.MousePointer = vbDefault
End Sub

Private Sub PadronizaRelatorio(objRelatorio As Object, adoRelatorio As ADODB.Recordset, tdbGridRelatorio As TDBGrid)
Dim intContItemRelatorio    As Integer
Dim intContRecordset        As Integer
Dim intCaracterField        As Integer
    
    intContItemRelatorio = -1
    
    For intContRecordset = 0 To adoRelatorio.Fields.Count - 1
        If tdbGridRelatorio.Columns(intContRecordset).Visible = True Then
            intContItemRelatorio = intContItemRelatorio + 1
            'Criação dos Labels
            objRelatorio.PageHeader.Controls.Add ("DDActiveReports2.Label")
            objRelatorio.PageHeader.Controls.Item("Label" & intContItemRelatorio + 1).Top = 1134
            objRelatorio.PageHeader.Controls.Item("Label" & intContItemRelatorio + 1).Left = 0
            objRelatorio.PageHeader.Controls.Item("Label" & intContItemRelatorio + 1).Width = Len(tdbGridRelatorio.Columns(intContRecordset).Caption) * 290
            objRelatorio.PageHeader.Controls.Item("Label" & intContItemRelatorio + 1).Caption = tdbGridRelatorio.Columns(intContRecordset).Caption
            'Criação dos Fields
            objRelatorio.Detail.Controls.Add ("DDActiveReports2.Field")
            objRelatorio.Detail.Controls.Item("Field" & intContItemRelatorio + 1).Top = 0
            objRelatorio.Detail.Controls.Item("Field" & intContItemRelatorio + 1).Style = "font-size: 8.5pt;"
            objRelatorio.Detail.Controls.Item("Field" & intContItemRelatorio + 1).DataField = tdbGridRelatorio.Columns(intContRecordset).DataField
            'Define tamanho dos fields
             
            tdbGridRelatorio.MoveFirst
            intCaracterField = 0
            Do While Not tdbGridRelatorio.EOF
                If Len(tdbGridRelatorio.Columns(intContRecordset).Value) > intCaracterField Then
                    intCaracterField = Len(RTrim(LTrim(tdbGridRelatorio.Columns(intContRecordset).Value)))
                End If
                tdbGridRelatorio.MoveNext
            Loop
            
            If adoRelatorio.Fields(intContRecordset).Type = 135 Then 'Data
                objRelatorio.Detail.Controls.Item("Field" & intContItemRelatorio + 1).Width = 945 + 220
            ElseIf adoRelatorio.Fields(intContRecordset).Type = adCurrency Or adoRelatorio.Fields(intContRecordset).Type = adDecimal Or _
                adoRelatorio.Fields(intContRecordset).Type = adDouble Or adoRelatorio.Fields(intContRecordset).Type = adNumeric Or _
                adoRelatorio.Fields(intContRecordset).Type = adInteger Or adoRelatorio.Fields(intContRecordset).Type = adVarNumeric Or _
                adoRelatorio.Fields(intContRecordset).Type = adSmallInt Then  'Campo de Qualquer valor numéricoThen 'Campo de Qualquer valor numérico
                
                objRelatorio.Detail.Controls.Item("Field" & intContItemRelatorio + 1).OutputFormat = "#,##0.00"
                objRelatorio.Detail.Controls.Item("Field" & intContItemRelatorio + 1).Width = intCaracterField * 160 + 220
            Else 'Alfanumérico
                objRelatorio.Detail.Controls.Item("Field" & intContItemRelatorio + 1).Width = intCaracterField * 220 + 220
            End If
         
            If intContItemRelatorio + 1 = 1 Then
                'Posicionar os Fields
                objRelatorio.Detail.Controls.Item("Field" & intContItemRelatorio + 1).Left = 50 'Posicão inicial do primeiro Field
                'Posicionar os Labels
                objRelatorio.PageHeader.Controls.Item("Label" & intContItemRelatorio + 1).Left = 50 'Posicão Inicial do primeiro Label
            Else
                'Posicionar os Fields
                objRelatorio.Detail.Controls.Item("Field" & intContItemRelatorio + 1).Left = objRelatorio.Detail.Controls.Item("Field" & intContItemRelatorio).Left + objRelatorio.Detail.Controls.Item("Field" & intContItemRelatorio).Width + 220
                'Posicionar os Labels
                objRelatorio.PageHeader.Controls.Item("Label" & intContItemRelatorio + 1).Left = objRelatorio.Detail.Controls.Item("Field" & intContItemRelatorio + 1).Left
            End If
        End If
    Next
    

End Sub

Public Sub gOrdenaGrid(objGrid As TDBGrid, IndiceDaColuna As Integer)

Dim recGrid           As ADODB.Recordset
Dim adoField          As ADODB.Field
Dim strField          As String
Dim i As Integer

    Set recGrid = objGrid.DataSource
    
    If Not recGrid Is Nothing Then
    
        strField = Space$(0)
        For i = 0 To recGrid.Fields.Count - 1
            If UCase(recGrid.Fields(i).Name) = UCase(objGrid.Columns(IndiceDaColuna).DataField) Then
                strField = objGrid.Columns(IndiceDaColuna).DataField
            End If
        Next
        
        If strField = Space$(0) Then Exit Sub
        
        gblnOrdenacaoAscGrid = IIf(gbytOrdenacaoGrid = IndiceDaColuna, Not gblnOrdenacaoAscGrid, False)
        gbytOrdenacaoGrid = IndiceDaColuna
        
        recGrid.Sort = objGrid.Columns(IndiceDaColuna).DataField & IIf(gblnOrdenacaoAscGrid, "  DESC", " ASC")
    
        If Not recGrid.EOF Then
            recGrid.MoveFirst
        End If
            
        Set objGrid.DataSource = recGrid
        objGrid.ReBind
        objGrid.Refresh
        
    End If
        
End Sub
Public Function GeraMovimentosByEvento(intEvento As Long, DTMDATA As String, dblValor As String, STRHISTORICO As String, strLancamento As String, strOrigem As String, Optional aryContas As Variant, Optional aryTpMov As Variant, Optional blnFlag As Boolean = False, Optional aryValor As Variant, Optional blnSemComit As Boolean, Optional blnCancelado As Boolean = False, Optional blnRepasseDecendial As Byte = 0) As Boolean
Dim strSql       As String
Dim adoResultado As New ADODB.Recordset
Dim intContador  As Integer
Dim lngPkidProcesso As Long

   
   'ATENÇÃO : PARA O CAMPO strOrigem teremos :
   '0 - Lançamentos Extras
   '1 - Tela de Programa de Trabalho
   '2 - Tela de Previsão da Receita
   '3 - Tela de Empenho
   '4 - Tela de Credito e Redução
   '5 - Tela de Remanejamento de Dotação
   '6 - Tela de Arrecadação da Receita
   '7 - Tela de Pagamento
   '8 - Tela de Transferência Bancária
   '9 - Tela de Adiantamentos
   '10- Tela de Carta Fiança
   '11- Adiantamento
   '12- Cancelamento de Restos a Pagar
   'Os parametros aryContas e aryTpMov devem começar com o indice 1
   
   
   GeraMovimentosByEvento = True
   
   dblValor = Replace(dblValor, ".", ",")
   'Pega todas as contas pertencentes ao evento selecionado
   
   strSql = "SELECT " 'Contas de Credito
       strSql = strSql & "EVC.intContaContabil, 0 AS TpMov , "
       strSql = strSql & gstrConvVrParaSql(dblValor) & " AS DBLVALOR, "
       strSql = strSql & "PC.blnPatrimonial, "
       strSql = strSql & "PC.bytMovimentaSistema "
   strSql = strSql & "FROM "
       strSql = strSql & gstrEventoContaContabilCredito & " EVC, "
       strSql = strSql & gstrPlanoConta & " PC "
   
   strSql = strSql & " WHERE "
      strSql = strSql & "EVC.intEvento = " & IIf(blnFlag = False, intEvento, 0)
      strSql = strSql & " AND EVC.intContaContabil = PC.Pkid"
      strSql = strSql & " AND EVC.bytContaGrupo = 0"
   
   strSql = strSql & " UNION ALL "
   
   strSql = strSql & "SELECT " 'Contas de Debito
       strSql = strSql & "EVD.intContaContabil, 1 AS TpMov , "
       strSql = strSql & gstrConvVrParaSql(dblValor) & " AS DBLVALOR, "
       strSql = strSql & "PC.blnPatrimonial, "
       strSql = strSql & "PC.bytMovimentaSistema "
       
   strSql = strSql & "FROM "
       strSql = strSql & gstrEventoContaContabilDebito & " EVD, "
       strSql = strSql & gstrPlanoConta & " PC "
   strSql = strSql & " WHERE "
       strSql = strSql & "EVD.intEvento = " & IIf(blnFlag = False, intEvento, 0)
       strSql = strSql & " AND EVD.intContaContabil = PC.Pkid"
       strSql = strSql & " AND EVD.bytContaGrupo = 0"
      
   'Vamos acrescentar as contas do array aryContas para serem gravadas junto das contas do evento
   If Not IsMissing(aryContas) Then
      For intContador = 1 To UBound(aryContas)
          strSql = strSql & " UNION ALL SELECT "
          If (bytDBType = EDatabases.SQLServer) Then
             strSql = strSql & " TOP 1 "
          End If
          strSql = strSql & aryContas(intContador) & " AS intContaContabil, "
          strSql = strSql & aryTpMov(intContador) & " AS TpMov, "
          If IsMissing(aryValor) Then
            strSql = strSql & gstrConvVrParaSql(dblValor) & " AS DBLVALOR, "
          Else
            strSql = strSql & gstrConvVrParaSql(aryValor(intContador)) & " AS DBLVALOR, "
          End If
          strSql = strSql & "PC.blnPatrimonial, "
          strSql = strSql & "PC.bytMovimentaSistema "
          
          strSql = strSql & " FROM " & gstrEvento & ", " & gstrPlanoConta & " PC "
          If (bytDBType = EDatabases.Oracle) Then
             strSql = strSql & " WHERE ROWNUM = 1 "
             strSql = strSql & " AND PC.Pkid = " & aryContas(intContador)
          End If
      Next
   End If
   'Esta rotina abaixo grava todas as contas pertencentes ao evento
   'com seus respectivos valores e tipos de saldo
      
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSql, 15, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            'Tabela Processo Pagamento
            Set gobjBanco = New clsBanco
            If Not blnSemComit Then gobjBanco.ExecutaBeginTrans
                        
            strSql = "INSERT INTO " & gstrProcessoPagamento & "("
            strSql = strSql & "intLancamentoContabil, dtmData, bytSituacao, bytNormal,"
            strSql = strSql & "strHistorico, intEvento, dtmDtAtualizacao, lngCodUsr, intLancamento, intOrigem, blnRepasseDecendial) "
            strSql = strSql & "SELECT " & gstrISNULL("MAX(intLancamentoContabil)", "0") & " + 1, "
            strSql = strSql & gstrConvDtParaSql(DTMDATA) & ", 0, 1, '" & STRHISTORICO & "', "
            If strOrigem = "8" Then
                strSql = strSql & intEvento & ","
            Else
                strSql = strSql & IIf(blnFlag = False, intEvento, "Null") & ", "
            End If
            strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strSql = strSql & glngCodUsr & ", " & strLancamento & ", " & strOrigem & ", " & blnRepasseDecendial & " FROM " & gstrProcessoPagamento
            
            Set gobjBanco = New clsBanco
            If gobjBanco.Execute(strSql) Then
                Set gobjBanco = New clsBanco
                If Not blnSemComit Then gobjBanco.ExecutaCommitTrans
            Else
                Set gobjBanco = New clsBanco
                If Not blnSemComit Then gobjBanco.ExecutaRollbackTrans
                GeraMovimentosByEvento = False
                Exit Function
            End If
            
            lngPkidProcesso = glngRetornaPkidTabelaPai("seq" & gstrProcessoPagamento, gstrProcessoPagamento)
            
            strSql = IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    
            
            While Not .EOF
               'Tabela Lancamento Contábil
               strSql = strSql & "INSERT INTO " & gstrLancamentoContabil & " ("
               strSql = strSql & "intProcesso, intConta, dblValor, bytNatureza, bytTipo, "
               strSql = strSql & "strDocumento, dtmDtAtualizacao, lngCodUsr) "
               strSql = strSql & "VALUES( " & lngPkidProcesso & ", " & !intContaContabil & ", "
               strSql = strSql & gstrConvVrParaSql(!dblValor) & ", " & !TpMov & " ,"
               strSql = strSql & IIf((blnCancelado = False), 0, 3) & ", '', "
               strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
               strSql = strSql & glngCodUsr & "); "
           
               
               ' GravaMovimentosSistemas
               If !blnPatrimonial = 1 And !bytMovimentaSistema = 0 Then
                  If strOrigem = "3" Then
                      If Not GeraMovimentosSistemas(!intContaContabil, !TpMov, !dblValor, CStr(IIf(IsNull(STRHISTORICO), "", STRHISTORICO)), DTMDATA, IIf((blnCancelado = False), 0, 3), BuscaFuncao(Val(strLancamento))) Then
                            Set gobjBanco = New clsBanco
                            If Not blnSemComit Then gobjBanco.ExecutaRollbackTrans
                            GeraMovimentosByEvento = False
                            Exit Function
                      End If
                   Else
                      If Not GeraMovimentosSistemas(!intContaContabil, !TpMov, !dblValor, CStr(IIf(IsNull(STRHISTORICO), "", STRHISTORICO)), DTMDATA, IIf((blnCancelado = False), 0, 3)) Then
                            Set gobjBanco = New clsBanco
                            If Not blnSemComit Then gobjBanco.ExecutaRollbackTrans
                            GeraMovimentosByEvento = False
                            Exit Function
                      End If
                   End If
               End If
                              
               .MoveNext
            
            Wend
            strSql = strSql & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
            
            Set gobjBanco = New clsBanco
            If gobjBanco.Execute(strSql) Then
                Set gobjBanco = New clsBanco
                If Not blnSemComit Then gobjBanco.ExecutaCommitTrans
            Else
                Set gobjBanco = New clsBanco
                If Not blnSemComit Then gobjBanco.ExecutaRollbackTrans
                GeraMovimentosByEvento = False
            End If
         Else
            GeraMovimentosByEvento = False
            Exit Function
         End If
      End With
   End If
   
   GeraMovimentosByEvento = True
   
End Function

Public Function GeraMovimentosSistemas(lngConta As Long, bytTipoLancamento As Byte, dblValor As String, STRHISTORICO As String, strData As String, bytTipoMovimento As Byte, Optional strFuncao As String, Optional intProcPagto As Integer) As Boolean
Dim strSql            As String
Dim adoResultado      As New ADODB.Recordset
Dim lngPkidConta      As Long
Dim gcmdADOCmdConMain As ADODB.Command
        
        GeraMovimentosSistemas = False
        
       dblValor = Replace(dblValor, ",", ".")
    
       strSql = "SELECT CZ.intPlanoContaDestino, bytTipoContaDestino, "
    
       If bytDBType = EDatabases.Oracle Then
          strSql = strSql & " DECODE(PC.blnFinanceira,1,1,DECODE(PC.bytVariacaoPatrimonial,1,2,DECODE(PC.blnOrcamentario,1,3,DECODE(PC.blnPatrimonial,1,4))) ) intSistema "
       Else
            'Alterado em 05/07/2004 por Wagner
            'Tive que fazer manualmente, pois a função gstrCaseWhen não estava servindo
            'Para o propósito do código (dava erro de array)
            'strSql = strSql & gstrCASEWHEN("PC.blnFinanceira", "1,1," & gstrCASEWHEN("PC.bytVariacaoPatrimonial", "1,2," & gstrCASEWHEN("PC.blnOrcamentario", "1,3," & gstrCASEWHEN("PC.blnPatrimonial", "1,4")))) & " intSistema "
            strSql = strSql & "(CASE  PC.blnFinanceira " _
                                & " WHEN 1 THEN 1 " _
                                & " Else: Case PC.bytVariacaoPatrimonial " _
                                & "         WHEN 1 THEN 2 " _
                                & "         Else: Case PC.blnOrcamentario " _
                                & "                 WHEN 1 THEN 3 " _
                                & "                 Else: Case PC.blnPatrimonial " _
                                & "                         WHEN  1 THEN 4 " _
                                & "                         End " _
                                & "                 End " _
                                & "         End " _
                                & " END) intSistema "
       End If
       strSql = strSql & " FROM " & gstrPlanoConta & " PC,"
       strSql = strSql & gstrCruzamentos & " CZ"
       strSql = strSql & " WHERE CZ.intPlanoContaDestino = PC.PKID AND CZ.intPlanoContaOrigem = " & lngConta
       strSql = strSql & " AND CZ.bytTipoLancamento = " & bytTipoLancamento
       strSql = strSql & " AND CZ.bytTipoMovimento = " & bytTipoMovimento
    
       gobjBanco.CriaADO strSql, 5, adoResultado
    
       With adoResultado
          If Not .EOF Then
    
             strSql = IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    
             strSql = strSql & "INSERT INTO " & gstrMovimentoSistemas
             strSql = strSql & "(intNumero, dtmData, strHistorico, dtmDtAtualizacao, lngCodUsr"
             If intProcPagto > 0 Then
                strSql = strSql & "," & "intProcesso"
              End If
              strSql = strSql & " ) "
             strSql = strSql & "SELECT " & gstrISNULL("MAX(intNumero)", "0", "MAX(intNumero)") & " + 1,"
             strSql = strSql & gstrConvDtParaSql(strData) & ", '" & STRHISTORICO & "', "
             strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr
             If intProcPagto > 0 Then
                strSql = strSql & ", " & intProcPagto
             End If
             strSql = strSql & " FROM " & gstrMovimentoSistemas & "; "
    
    
             While Not .EOF
    
                If !bytTipoContaDestino = 1 Then
    
                   If Len(strFuncao) > 0 And BuscaContaContabil(!intPlanoContaDestino, strFuncao) > 0 Then
    
                      lngPkidConta = BuscaContaContabil(!intPlanoContaDestino, strFuncao)
    
                   Else
    
                      'ExibeMensagem "Não foi possível gravar este movimento para os demais sistemas."
    
                      strSql = "INSERT INTO " & gstrMovimentoSistemas
                      strSql = strSql & "(intNumero, dtmData, strHistorico, dtmDtAtualizacao, lngCodUsr "
                      If intProcPagto > 0 Then
                        strSql = strSql & "," & "intProcesso"
                      End If
                      strSql = strSql & " ) "
                      strSql = strSql & "SELECT " & gstrISNULL("MAX(intNumero)", "0", "MAX(intNumero)") & " + 1, "
                      strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                                      'CONTA DE ORIGEM * 1 debito e 0 credito * 0 normal e 3 cancelamento * VALOR * FUNCAO
                      strSql = strSql & "'*" & lngConta & "*" & bytTipoLancamento & "*" & bytTipoMovimento & "*" & dblValor & "*" & strFuncao & "', "
                      strSql = strSql & gstrConvDtParaSql(strData) & ", " & glngCodUsr
                      If intProcPagto > 0 Then
                        strSql = strSql & ", " & intProcPagto
                      End If
                      strSql = strSql & " FROM " & gstrMovimentoSistemas
    
                      If Not gobjBanco.Execute(strSql) Then
                        GeraMovimentosSistemas = False
                      Else
                        GeraMovimentosSistemas = True
                      End If
                      
                      Exit Function
    
                   End If
    
                Else
    
                   lngPkidConta = !intPlanoContaDestino
    
                End If
                    
                strSql = strSql & "INSERT INTO " & gstrContaMovimentoSistemas
                strSql = strSql & "(intMovimentoSistema, intSistema, intPlanoConta, bytTipo, dblValor, dtmDtAtualizacao, lngCodUsr) "
                strSql = strSql & "SELECT MAX(PKID), " & !intSistema & ", " & lngPkidConta & ", "
                strSql = strSql & bytTipoLancamento & ", (" & (dblValor) & "), "
                strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr
                strSql = strSql & " FROM " & gstrMovimentoSistemas & ";"
    
                .MoveNext
    
             Wend
    
             strSql = strSql & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
    
             Set gobjBanco = New clsBanco
             If Not gobjBanco.Execute(strSql) Then
                Exit Function
             Else
                GeraMovimentosSistemas = True
             End If
    
          Else
    
             'ExibeMensagem "Não há sistema para este tipo de movimento."
    
             strSql = "INSERT INTO " & gstrMovimentoSistemas
             strSql = strSql & "(intNumero, dtmData, strHistorico, dtmDtAtualizacao, lngCodUsr "
             If intProcPagto > 0 Then
                strSql = strSql & "," & "intProcesso"
             End If
             strSql = strSql & " ) "
             strSql = strSql & "SELECT " & gstrISNULL("MAX(intNumero)", "0") & " + 1, "
             strSql = strSql & gstrConvDtParaSql(strData) & ", "
                             'CONTA DE ORIGEM * 1 debito e 0 credito * 0 normal e 3 cancelamento * VALOR * FUNCAO
             strSql = strSql & "'*" & lngConta & "*" & bytTipoLancamento & "*" & bytTipoMovimento & "*" & dblValor & "*" & strFuncao & "', "
             strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr
             If intProcPagto > 0 Then
                strSql = strSql & ", " & intProcPagto
             End If
             strSql = strSql & " FROM " & gstrMovimentoSistemas
    
             If Not gobjBanco.Execute(strSql) Then
                Exit Function
             Else
                GeraMovimentosSistemas = True
             End If
         End If
    End With
    
    GeraMovimentosSistemas = True
    
End Function

Public Function BuscaFuncao(lngLancamento As Long) As String
   Dim strSql     As String
   Dim adoResultado   As New ADODB.Recordset

   strSql = "SELECT FG.strCodigo FROM "
   strSql = strSql & gstrFuncaoDoGoverno & " FG,"
   strSql = strSql & gstrEmpenho & " EP,"
   strSql = strSql & gstrProgramaDeTrabalho & " PT "
   strSql = strSql & "WHERE EP.intProgramaTrabalho = PT.PKID "
   strSql = strSql & "AND PT.intFuncao = FG.PKID "
   'obs. os exercicios foram alterados por gintExercicio 2 ln
   strSql = strSql & "AND PT.intExercicio = " & gintExercicio & " "
   strSql = strSql & "AND EP.intNumero =  " & lngLancamento & " AND " & gstrDATEPART("yyyy", "EP.dtmData") & " = " & gintExercicio

   gobjBanco.CriaADO strSql, 5, adoResultado

   With adoResultado
      If Not .EOF Then
         BuscaFuncao = !strCodigo
      End If
   End With

End Function

Public Function BuscaContaContabil(intContaDestino As Long, strFuncao As String) As Long
   Dim strSql            As String
   Dim adoResultado      As New ADODB.Recordset
   Dim strCodigoContabil As String

   strCodigoContabil = Replace(gstrMascaraContaContabil, ".", "")

   strSql = "SELECT " & strSUBSTRING & "(strContaContabil,1," & Len(strCodigoContabil) - Len(strFuncao) & ")" & strCONCAT & " '" & strFuncao & "'" & " strCodigo "
   strSql = strSql & "FROM " & gstrPlanoConta & " WHERE PKID = " & intContaDestino

   gobjBanco.CriaADO strSql, 5, adoResultado

   If Not adoResultado.EOF Then
      strSql = "SELECT PKID FROM " & gstrPlanoConta & " WHERE strContaContabil LIKE '" & adoResultado!strCodigo & "%'"
   Else
      BuscaContaContabil = 0
      Exit Function
   End If

   gobjBanco.CriaADO strSql, 5, adoResultado

   If Not adoResultado.EOF Then
      BuscaContaContabil = adoResultado!Pkid
   Else
      BuscaContaContabil = 0
      Exit Function
   End If

End Function

Public Function gstrCalculaDigitoModulo10(strCodigo As String) As String
        Dim intTam      As Integer
        Dim intDv       As Integer
        Dim intMULT     As Integer
        Dim intResult   As Integer
        Dim intSoma     As Integer
        Dim intResto    As Integer
        Dim blnV_Resp   As Boolean
        Dim strChar2    As String

    intTam = Len(strCodigo)
    intMULT = 1
    
    blnV_Resp = True
    
    Do While intTam >= 1
        intMULT = intMULT + 1
        If intMULT > 2 Then
            intMULT = 1
        End If
        intResult = Mid(strCodigo, intTam, 1) * intMULT
        
        If intResult >= 10 Then
            strChar2 = intResult
            intSoma = intSoma + Val(Mid(strChar2, 1, 1)) + Val(Mid(strChar2, 2, 1))
        Else
            intSoma = intSoma + intResult
        End If
        intTam = intTam - 1
    
    Loop
    intResto = intSoma Mod 10
    If intResto = 0 Then
        intDv = 0
    Else
        intDv = 10 - intResto
    End If
    
    gstrCalculaDigitoModulo10 = intDv

End Function

Public Function gstrCalculaDigitoAutoConferencia(strCodigo As String) As String
        Dim intTam      As Integer
        Dim intDv       As Integer
        Dim intMULT     As Integer
        Dim intResult   As Integer
        Dim intSoma     As Integer
        Dim intResto    As Integer
        Dim blnV_Resp   As Boolean
        Dim strChar2    As String

    intTam = Len(strCodigo)
    intMULT = 1
    
    blnV_Resp = True
    
    Do While intTam >= 1
        intMULT = intMULT + 1
        If intMULT > 9 Then
            intMULT = 2
        End If
        intResult = Mid(strCodigo, intTam, 1) * intMULT
        
        intSoma = intSoma + intResult

        intTam = intTam - 1
    
    Loop
    intResto = intSoma Mod 11
    
    If intResto = 0 Or intResto = 1 Or intResto > 9 Then
        intDv = 1
    Else
        intDv = 11 - intResto
    End If
    
    gstrCalculaDigitoAutoConferencia = intDv

End Function

Public Function gstrCalculaDigitoNossoNumero(strCodigo As String) As String
        Dim intTam      As Integer
        Dim intDv       As Integer
        Dim intMULT     As Integer
        Dim intResult   As Integer
        Dim intSoma     As Integer
        Dim intResto    As Integer
        Dim blnV_Resp   As Boolean
        Dim strChar2    As String

    intTam = Len(strCodigo)
    intMULT = 1
    
    blnV_Resp = True
    
    Do While intTam >= 1
        
        If intMULT = 1 Then
            intMULT = 3
        ElseIf intMULT = 3 Then
            intMULT = 7
        ElseIf intMULT = 7 Then
            intMULT = 9
        Else
            intMULT = 1
        End If
        intResult = Mid(strCodigo, intTam, 1) * intMULT
        If intResult >= 10 Then
            strChar2 = intResult
            intSoma = intSoma + Val(Mid(strChar2, 2, 1))
        Else
            intSoma = intSoma + intResult
        End If
        
        intTam = intTam - 1
    
    Loop
    
    intResto = Right(Str(intSoma), 1)
    
    intDv = 10 - intResto

    intDv = Right(Str(intDv), 1)
    
    gstrCalculaDigitoNossoNumero = intDv

End Function

Public Function gstrCalculaDigito2Asbace(strCodigo As String) As String
        Dim intTam      As Integer
        Dim strDv       As String
        Dim intMULT     As Integer
        Dim intResult   As Integer
        Dim intSoma     As Integer
        Dim intResto    As Integer
        Dim blnV_Resp   As Boolean
        Dim strChar2    As String

Recalcular:

    intSoma = 0
    
    intTam = Len(strCodigo)

    intMULT = 1
    
    blnV_Resp = True
    
    Do While intTam >= 1
        intMULT = intMULT + 1
        If intMULT > 7 Then
            intMULT = 2
        End If
        intResult = Mid(strCodigo, intTam, 1) * intMULT
        


        intSoma = intSoma + intResult

        intTam = intTam - 1
    
    Loop

    intResto = intSoma Mod 11

    
    'Sempre será retornado o digito1 e digito2, pelo fato do digito1 poder ser alterado
    If intResto = 0 Then
        strDv = Mid(strCodigo, Len(strCodigo), 1) & "0"
    ElseIf intResto = 1 Then

        'Caso o resto seja 1, vamos recalcular modificando o digito1, até o resto ser diferente de 1
        
        If Right(strCodigo, 1) = 9 Then
            strCodigo = Mid(strCodigo, 1, Len(strCodigo) - 1) & 0
        Else
            strCodigo = Mid(strCodigo, 1, Len(strCodigo) - 1) & (Right(strCodigo, 1) + 1)
        End If

        GoTo Recalcular

    Else
        strDv = Mid(strCodigo, Len(strCodigo), 1) & 11 - intResto
    End If
    
    gstrCalculaDigito2Asbace = strDv

End Function

Public Function VerificaItemNoAlmoxarifado(dbcItem As DataCombo, dbcAlmoxarifado As DataCombo) As Boolean

Dim strSql As String
Dim adoRec As ADODB.Recordset

VerificaItemNoAlmoxarifado = False

    If IsNull(dbcItem.BoundText) = False Then
    
        strSql = "SELECT PkId FROM " & gstrMaterialEmEstoque
        strSql = strSql & " WHERE intCatalogoMaterial = " & dbcItem.BoundText
        strSql = strSql & " AND intAlmoxarifado = " & dbcAlmoxarifado.BoundText
        
        Set gobjBanco = New clsBanco
        
        If gobjBanco.CriaADO(strSql, 10, adoRec) Then
                
            If adoRec.EOF And adoRec.BOF Then
                Exit Function
            End If
                       
        End If
    
    End If
    
VerificaItemNoAlmoxarifado = True

End Function

Public Function VerificaGrupoMaterial(dbcCodigo As DataCombo, dbcDescricao As DataCombo) As Boolean

Dim strSql As String
Dim adoRec As ADODB.Recordset

    VerificaGrupoMaterial = False
    
    If dbcCodigo.Text = "" Then
        Exit Function
    End If
    
    strSql = "SELECT PkId"
    strSql = strSql & " FROM " & gstrGrupoMaterialServico
    strSql = strSql & " WHERE intCodigo = " & dbcCodigo
    strSql = strSql & " ORDER BY intCodigo "
    
    Set gobjBanco = New clsBanco
        
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    
        With adoRec
    
            If Not .BOF And Not .EOF Then
            
                PreencherListaDeOpcoes dbcDescricao, !Pkid
                
                VerificaGrupoMaterial = True
                
            Else
            
                dbcCodigo.Text = ""
                dbcCodigo.BoundText = ""
                
                dbcDescricao.Text = ""
                dbcDescricao.BoundText = ""
                            
            End If
            
        End With
        
    End If
    
    Set gobjBanco = Nothing

End Function

Public Function VerificaTipoMaterial(dbcCodigo As DataCombo, dbcDescricao As DataCombo, dbcGrupoMaterialServico As DataCombo, Optional blnCatalogoItem As Boolean) As Boolean

Dim strSql As String
Dim adoRec As ADODB.Recordset

    VerificaTipoMaterial = False
    
    If dbcCodigo.Text = "" Then
        Exit Function
    End If
    
    strSql = "SELECT PkId"
    strSql = strSql & " FROM " & gstrTipoMaterialServico
    strSql = strSql & " WHERE intCodigo = " & dbcCodigo
    strSql = strSql & " AND intGrupoMaterialServico = " & dbcGrupoMaterialServico.BoundText
    If blnCatalogoItem Then
        strSql = strSql & " AND bytIdentificaTipo = 1 "
    End If
    strSql = strSql & " ORDER BY intCodigo "
    
    Set gobjBanco = New clsBanco
        
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    
        With adoRec
    
            If Not .BOF And Not .EOF Then
            
'                PreencherListaDeOpcoes dbcCodigo, !Pkid
                PreencherListaDeOpcoes dbcDescricao, !Pkid
'                dbcDescricao.HelpContextID = 1
'                LeDaTabelaParaObj "", dbcDescricao, "SELECT PkId, strDescricao FROM " & gstrTipoMaterialServico & " WHERE intCodigo = " & dbcCodigo.Text & " AND intGrupoMaterialServico = " & dbcGrupoMaterialServico.BoundText & " ORDER BY strDescricao "
'                dbcDescricao.HelpContextID = 0
                
                VerificaTipoMaterial = True
                
            Else
            
                dbcCodigo.Text = ""
                dbcCodigo.BoundText = ""
                
                dbcDescricao.Text = ""
                dbcDescricao.BoundText = ""
                            
            End If
            
        End With
        
    End If
    
    Set gobjBanco = Nothing

End Function

Public Function VerificaFamiliaMaterial(dbcCodigo As DataCombo, dbcDescricao As DataCombo, dbcTipoMaterialServico As DataCombo) As Boolean

Dim strSql As String
Dim adoRec As ADODB.Recordset

    VerificaFamiliaMaterial = False
    
    If dbcCodigo.Text = "" Then
        Exit Function
    End If
    
    strSql = "SELECT PkId"
    strSql = strSql & " FROM " & gstrFamiliaMaterial
    strSql = strSql & " WHERE strCodigo = " & dbcCodigo
    strSql = strSql & " AND intTipoMaterialServico = " & dbcTipoMaterialServico.BoundText
    strSql = strSql & " ORDER BY strCodigo "
    
    Set gobjBanco = New clsBanco
        
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    
        With adoRec
    
            If Not .BOF And Not .EOF Then
            
'                PreencherListaDeOpcoes dbcCodigo, !Pkid
                PreencherListaDeOpcoes dbcDescricao, !Pkid
'                dbcDescricao.HelpContextID = 1
'                LeDaTabelaParaObj "", dbcDescricao, "SELECT PkId, strDescricao FROM " & gstrTipoMaterialServico & " WHERE intCodigo = " & dbcCodigo.Text & " AND intGrupoMaterialServico = " & dbcGrupoMaterialServico.BoundText & " ORDER BY strDescricao "
'                dbcDescricao.HelpContextID = 0
                
                VerificaFamiliaMaterial = True
                
            Else
            
                dbcCodigo.Text = ""
                dbcCodigo.BoundText = ""
                
                dbcDescricao.Text = ""
                dbcDescricao.BoundText = ""
                            
            End If
            
        End With
        
    End If
    
    Set gobjBanco = Nothing

End Function

Public Function glngRetornaPkidTabelaPai(strNomeSequence As String, strNomeTabela As String) As Long
Dim strSql As String
Dim adoRec As ADODB.Recordset

'A instrução em SQL Server é a seguinte : SELECT IDENT_CURRENT('?') onde o ? você substitui pelo nome da tabela que você deseja ....
'A instrução do Oracle é a seguinte ... "SELECT seq?.CURRVAL FROM DUAL" onde seq? seria o nome da sequencia que você criou para a tabela digo ....

    If bytDBType = EDatabases.Oracle Then
        strSql = "SELECT " & strNomeSequence & ".CURRVAL Pkid FROM Dual"
    Else
        strSql = "SELECT IDENT_CURRENT('" & strNomeTabela & "') Pkid"
    End If
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    
        With adoRec
    
            If Not .BOF And Not .EOF Then
                glngRetornaPkidTabelaPai = adoRec("Pkid").Value
            End If
            
        End With
        
    End If

End Function

Public Function LeCDCCredor(Optional strPKId As String, Optional strCDC As String)

Dim strSql          As String
Dim adoResultado    As ADODB.Recordset
    
    If Trim(strPKId) = "" And Trim(strCDC) = "" Then
        LeCDCCredor = ""
        Exit Function
    End If
    
    strSql = ""
    strSql = strSql & "SELECT CDC , PKID"
    strSql = strSql & " FROM "
    strSql = strSql & gstrContribuinte
    If strPKId <> "" Then
        strSql = strSql & " WHERE PKID = " & strPKId
    ElseIf strCDC <> "" Then
        strSql = strSql & " WHERE CDC = " & strCDC
    End If
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            If strPKId <> "" Then
                LeCDCCredor = gstrENulo(adoResultado!CDC)
            ElseIf strCDC <> "" Then
                LeCDCCredor = gstrENulo(adoResultado!Pkid)
            End If
        End If
        
    End If
End Function

Public Function LeCoditemDespesa(Optional strPKId As String, Optional strCod As String) As String

Dim strSql          As String
Dim adoResultado    As ADODB.Recordset
    
    If Trim(strPKId) = "" And Trim(strCod) = "" Then
        LeCoditemDespesa = ""
        Exit Function
    End If
    
    strSql = ""
    strSql = strSql & "SELECT strcodigo , PKID"
    strSql = strSql & " FROM "
    strSql = strSql & gstrItemDespesa
    If strPKId <> "" Then
        strSql = strSql & " WHERE PKID = " & strPKId
    ElseIf strCod <> "" Then
        strSql = strSql & " WHERE strCodigo LIKE '%" & Replace(strCod, ".", "") & "%'"
    End If
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            If strPKId <> "" Then
                LeCoditemDespesa = gvntFormatacaoEspecifica(gstrENulo(adoResultado!strCodigo), 4)
            ElseIf strCod <> "" Then
                LeCoditemDespesa = gstrENulo(adoResultado!Pkid)
            End If
        End If
        
    End If
End Function



Public Sub CarregaTamanhoMascaras()
Dim strSql As String
Dim adoResultado As New ADODB.Recordset

    strSql = ""
    strSql = strSql & "SELECT intTipoDeInscricao, SUM(intTamanho) intTamanho FROM " & gstrCampoDeInscricao & " "
    strSql = strSql & "GROUP BY intTipoDeInscricao "
    strSql = strSql & "ORDER BY intTipoDeInscricao"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                
                Select Case adoResultado("intTipoDeInscricao")
                    
                    Case Is = TYP_IMOBILIARIA
                        vetTamanhoMascaras.intMaskImobiliario = adoResultado("intTamanho").Value
                    Case Is = TYP_ECONOMICA
                        vetTamanhoMascaras.intMaskEconomico = adoResultado("intTamanho").Value
                    Case Is = TYP_DIVIDA_ATIVA
                        vetTamanhoMascaras.intMaskDividaAtiva = adoResultado("intTamanho").Value
                    Case Is = TYP_ACORDO
                        vetTamanhoMascaras.intMaskAcordo = adoResultado("intTamanho").Value
                    Case Is = TYP_PRECO_PUBLICO
                        vetTamanhoMascaras.intMaskPrecoPublico = adoResultado("intTamanho").Value
                    Case Is = TYP_ISS_CONSTRUCAO
                        vetTamanhoMascaras.intMaskIssConstrucao = adoResultado("intTamanho").Value
                
                End Select
                
                .MoveNext
                
            Loop
        End With
    End If
    
End Sub

Public Function gintRetornaTamanhoMascara(bytTipoDeInscricao As Byte) As Integer
    
    Select Case bytTipoDeInscricao
        
        Case Is = TYP_IMOBILIARIA
            gintRetornaTamanhoMascara = vetTamanhoMascaras.intMaskImobiliario
        Case Is = TYP_ECONOMICA
            gintRetornaTamanhoMascara = vetTamanhoMascaras.intMaskEconomico
        Case Is = TYP_DIVIDA_ATIVA
            gintRetornaTamanhoMascara = vetTamanhoMascaras.intMaskDividaAtiva
        Case Is = TYP_ACORDO
            gintRetornaTamanhoMascara = vetTamanhoMascaras.intMaskAcordo
        Case Is = TYP_PRECO_PUBLICO
            gintRetornaTamanhoMascara = vetTamanhoMascaras.intMaskPrecoPublico
        Case Is = TYP_ISS_CONSTRUCAO
            gintRetornaTamanhoMascara = vetTamanhoMascaras.intMaskIssConstrucao
    
    End Select
    
    If gintRetornaTamanhoMascara = 0 Then gintRetornaTamanhoMascara = 20
    
End Function

Public Function gstrFormataInscricao(strInscricao As String, Optional intUtilizacao As Integer = 1) As String
Dim cont As Integer
Dim strSql As String
Dim adoRec As New ADODB.Recordset

    strSql = "SELECT intTamanho, strSeparador "
    strSql = strSql & "FROM " & gstrCampoDeInscricao
    strSql = strSql & " WHERE intTipodeinscricao = " & intUtilizacao
    strSql = strSql & "ORDER BY intSequencia "
    
    Set gobjBanco = New clsBanco
    
    cont = 1
    
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
        With adoRec
            If Not .EOF Then
                Do While Not .EOF
                    gstrFormataInscricao = gstrFormataInscricao & Mid(strInscricao, cont, !intTamanho) & IIf(Trim(gstrENulo(!strSeparador)) <> "", Trim(gstrENulo(!strSeparador)), "")
                    cont = cont + !intTamanho
                    .MoveNext
                Loop
            Else
                gstrFormataInscricao = strInscricao
            End If
        End With
    End If
    
End Function


Public Function VerificaDataEncerramento(strSistema As String, intExercicio As Integer) As Date

   Dim strSql As String
   Dim adoResultado As ADODB.Recordset
   
   strSql = "SELECT dtmFechamento FROM " & gstrFechamentoContabil
   strSql = strSql & " WHERE strCodigo = '" & strSistema & "'"
   strSql = strSql & " AND intExercicio = " & intExercicio
   
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 15, adoResultado) Then
       If Not adoResultado.EOF Then
          VerificaDataEncerramento = CDate(gstrDataFormatada(adoResultado!dtmFechamento))
       Else
          ExibeMensagem "Nenhum encerramento foi encontrado para este exercício."
          VerificaDataEncerramento = Empty
       End If
    End If
    
    adoResultado.Close: Set adoResultado = Nothing
    Set gobjBanco = Nothing
    
End Function


Public Function RetornaSituacaoExercicio(intExercicio As Integer) As String

Dim strSql As String
Dim adoTemp As New ADODB.Recordset

strSql = " SELECT bytSituacao FROM " & gstrExercicio & " WHERE intExercicio = " & intExercicio

Set gobjBanco = New clsBanco

If gobjBanco.CriaADO(strSql, 10, adoTemp) Then
    If Not adoTemp.EOF Then
        Select Case adoTemp!bytSituacao
            Case 0
                RetornaSituacaoExercicio = "Proposta"
            Case 1
                RetornaSituacaoExercicio = "Aberto"
            Case 2
                RetornaSituacaoExercicio = "Fechado"
            Case Else
                RetornaSituacaoExercicio = ""
        End Select
    End If
End If


End Function

Public Function blnValidarProcesso() As Boolean

   Dim strSql       As String
   Dim adoResultado As ADODB.Recordset

   strSql = "SELECT bytValidarProcesso FROM " & gstrConfiguracaoGeral
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
      
      If gstrENulo(adoResultado!bytValidarProcesso) > 0 Then
      
         blnValidarProcesso = True
      
      End If
   
   End If
End Function

Public Function gstrContaValoresAcumuladosEduc() As String

Dim strSql       As String
Dim adoResultado As ADODB.Recordset

   strSql = "SELECT bytOrcAteNivelModalidade FROM " & gstrConfiguracaoGeral
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
           If Not adoResultado.EOF Then
                gstrContaValoresAcumuladosEduc = IIf((adoResultado!bytOrcAteNivelModalidade = 1), "vw_contavalacumulados_educ_B", "vw_contavalacumulados_educ_A")
           End If
   End If
End Function

Public Sub ImprimePrecoPublico(strInscricao As String, _
                                        intNumeroDaGuia As Long, _
                                        strContribuinte As String, _
                                        strLogradouro As String, _
                                        STRBAIRRO As String, _
                                        STRMUNICIPIO As String, _
                                        STRUF As String, _
                                        INTCEP As Long, _
                                        strQuadra As String, _
                                        strLote As String, _
                                        strAviso As String, _
                                        strProcesso As String, _
                                        strsigla As String, _
                                        STRHISTORICO As String, _
                                        intConta As Long, _
                                        ByVal dblValor As Double, _
                                        ByVal dblCorrecao As Double, _
                                        ByVal dblMulta As Double, _
                                        ByVal dblJuros As Double, _
                                        ByVal dblTotal, _
                                        dtmDataVencimento As String, vetParecelas() As String, _
                                        Optional blnFebraban As Boolean = True)
                                        
Dim intFor          As Integer

Dim strCodBarras    As String
Dim adoResultado    As ADODB.Recordset
Dim strSql          As String

Dim lngGuias        As Long

Dim intFebraban     As Integer
Dim INTNUMERO       As Long
Dim bytDigito       As Byte
Dim strNumeroBoleto As String
Dim strNossoNumero  As String
               
Dim vetGuiaPrecoPublico() As String
    
    ReDim vetGuiaPrecoPublico(31, 0)
    
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    
    'Query utilizada para pegar o Codigo Febraban da tblEmpresa
    strSql = ""
    strSql = strSql & "Select * From " & gstrEmpresa
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 30, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            If gstrENulo(adoResultado!intFebraban) <> "" Then
                intFebraban = gstrENulo(adoResultado!intFebraban)
            Else
                ExibeMensagem "Código Febraban não encontrado."
                Exit Sub
            End If
        Else
            ExibeMensagem "Código Febraban não encontrado."
            Exit Sub
        End If
    End If
    
    'Número da Guia
    INTNUMERO = intNumeroDaGuia
    
    'Vamos definir o codigo de barras
    strCodBarras = gstrMontaCodigoBarras(IIf(blnFebraban, FEBRABAN, FICHA_COMPENSACAO), intConta, dblTotal, dtmDataVencimento, intFebraban, INTNUMERO, True, True)
    If Len(strCodBarras) = 0 Then Exit Sub
    'Vamos definir a linha digitavel
    strNumeroBoleto = gstrMontaLinhaDigitavel(IIf(blnFebraban, FEBRABAN, FICHA_COMPENSACAO), strCodBarras)
    'Vamos definir o nosso numero
    If Not blnFebraban Then
        strNossoNumero = gstrMontaNossoNumero(intConta, INTNUMERO)
    End If
    
    'Vamos inserir a guia na tabela TblGuias
    strSql = ""
    strSql = strSql & "Insert Into " & gstrGuias & "("
    strSql = strSql & "Intcontabancaria, "
    strSql = strSql & "Intnumero, "
    strSql = strSql & "Dtmdtemissao, "
    strSql = strSql & "Dblvalor, "
    strSql = strSql & "Strcodbarra, "
    strSql = strSql & "Dtmdtatualizacao, "
    strSql = strSql & "Lngcodusr, "
    strSql = strSql & "Dtmdtvencimento "
    strSql = strSql & ") Values("
    strSql = strSql & IIf(intConta = 0, "Null", intConta) & ", "
    strSql = strSql & INTNUMERO & ", "
    strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
    strSql = strSql & gstrConvVrParaSql(dblTotal) & ", '"
    strSql = strSql & strCodBarras & "', "
    strSql = strSql & strGETDATE & ", "
    strSql = strSql & glngCodUsr & ", "
    strSql = strSql & gstrConvDtParaSql(dtmDataVencimento)
    strSql = strSql & ")"
    
    If Not gobjBanco.Execute(strSql) Then
        ExibeMensagem "Erro na gravação da guia."
        gobjBanco.ExecutaRollbackTrans
        Exit Sub
    End If
    
    lngGuias = glngRetornaPkidTabelaPai("seqTblGuias", gstrGuias)
    
    'Vamos inserir as parcelas na tabela TblLancamentoGuias
    strSql = IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    
    For intFor = 0 To UBound(vetParecelas(), 2)
        strSql = strSql & "Insert Into " & gstrLancamentoGuias & "("
        strSql = strSql & "intlancamentovalor, "
        strSql = strSql & "intguias, "
        strSql = strSql & "dblvalorprincipal, "
        strSql = strSql & "dblvalormulta, "
        strSql = strSql & "dblvalorjuros, "
        strSql = strSql & "dblvalorcorrecao, "
        strSql = strSql & "dblvalordesconto, "
        strSql = strSql & "dtmdtatualizacao, "
        strSql = strSql & "lngcodusr) "
        strSql = strSql & "Values ("
        strSql = strSql & vetParecelas(0, intFor) & ", "
        strSql = strSql & lngGuias & ", "
        strSql = strSql & gstrConvVrParaSql(vetParecelas(1, intFor)) & ", "
        strSql = strSql & gstrConvVrParaSql(vetParecelas(2, intFor)) & ", "
        strSql = strSql & gstrConvVrParaSql(vetParecelas(3, intFor)) & ", "
        strSql = strSql & gstrConvVrParaSql(vetParecelas(4, intFor)) & ", "
        strSql = strSql & gstrConvVrParaSql("0") & ", "
        strSql = strSql & strGETDATE & ", "
        strSql = strSql & glngCodUsr & ") "
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), ";", "")
    Next
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
    
    If gobjBanco.Execute(strSql) Then
        gobjBanco.ExecutaCommitTrans
    Else
        ExibeMensagem "Erro na gravação da guia."
        gobjBanco.ExecutaRollbackTrans
        Exit Sub
    End If
                
    vetGuiaPrecoPublico(0, 0) = INTNUMERO & "/" & Year(gstrDataDoSistema)
    vetGuiaPrecoPublico(1, 0) = dtmDataVencimento
    vetGuiaPrecoPublico(2, 0) = strContribuinte
    vetGuiaPrecoPublico(3, 0) = strLogradouro
    vetGuiaPrecoPublico(4, 0) = STRBAIRRO
    vetGuiaPrecoPublico(5, 0) = strQuadra
    vetGuiaPrecoPublico(6, 0) = strLote
    vetGuiaPrecoPublico(7, 0) = strInscricao
    vetGuiaPrecoPublico(8, 0) = strAviso
    vetGuiaPrecoPublico(9, 0) = strsigla
    vetGuiaPrecoPublico(10, 0) = STRHISTORICO
    vetGuiaPrecoPublico(11, 0) = gstrConvVrDoSql(dblValor)
    vetGuiaPrecoPublico(12, 0) = gstrConvVrDoSql(dblCorrecao)
    vetGuiaPrecoPublico(13, 0) = gstrConvVrDoSql(dblMulta)
    vetGuiaPrecoPublico(14, 0) = gstrConvVrDoSql(dblJuros)
    vetGuiaPrecoPublico(15, 0) = gstrConvVrDoSql(dblTotal)
    vetGuiaPrecoPublico(16, 0) = gstrDataDoSistema
    vetGuiaPrecoPublico(17, 0) = gstrLoginUser
    vetGuiaPrecoPublico(18, 0) = dtmDataVencimento
    vetGuiaPrecoPublico(19, 0) = strNumeroBoleto
    vetGuiaPrecoPublico(20, 0) = strCodBarras
    vetGuiaPrecoPublico(21, 0) = STRMUNICIPIO
    vetGuiaPrecoPublico(22, 0) = STRUF
    vetGuiaPrecoPublico(23, 0) = strProcesso
    vetGuiaPrecoPublico(24, 0) = INTCEP
    vetGuiaPrecoPublico(25, 0) = intConta
    vetGuiaPrecoPublico(26, 0) = strNossoNumero
    
    'Query utilizada para pegar dados do Alfa
    strSql = ""
    strSql = strSql & "Select LV.intLancamentoAlfa, LV.intParcela, LV.bitParcelaValida, LA.strComposicaoDaReceita, LA.intExercicio " & _
                      " From " & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA " & _
                      " Where LV.Pkid = " & vetParecelas(0, 0) & " and LA.pkid = LV.intlancamentoalfa "
                      
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 30, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            vetGuiaPrecoPublico(27, 0) = gstrENulo(adoResultado!strComposicaoDaReceita)
            vetGuiaPrecoPublico(28, 0) = gstrENulo(adoResultado!intExercicio)
            vetGuiaPrecoPublico(29, 0) = gstrENulo(adoResultado!intParcela)
            vetGuiaPrecoPublico(30, 0) = gstrENulo(adoResultado!bitParcelaValida)
            vetGuiaPrecoPublico(31, 0) = gstrENulo(adoResultado!intLancamentoAlfa)
        Else
            ExibeMensagem "Lançamento Alfa não encontrado."
            Exit Sub
        End If
    End If
    
    'Vamos imprimir o relatorio de guia de arrecadacao
    If Not IsNull(vetGuiaPrecoPublico(0, 0)) Then
        'Vamos verificar se é febraban ou ficha compensacao
        If blnFebraban Then
            ImprimeRelatorioPorArray rptGuiaPrecoPublico, vetGuiaPrecoPublico, "Guia de Arrecadação"
        Else
            ImprimeRelatorioPorArray rptGuiaFichaPrecoPublico, vetGuiaPrecoPublico, "Guia de Arrecadação"
        End If
    End If
    
End Sub

Public Function glngRetornaProximoNumeroGuia(Optional strTabela As String = gstrParametrosTributario, Optional strCampo As String = "intNumeroGuia") As Long
Dim adoResultado           As New ADODB.Recordset
Dim adoCommand             As ADODB.Command
Dim strSql                 As String
Dim intRegistros           As Integer

On Error GoTo Problema_Na_Rotina

    'Vamos simular uma sequence para gerar o numero da guia
    Set gobjBanco = New clsBanco
    
ProximoNumeroGuia:

    'Vamos obter o numero da guia a ser atualizado
    strSql = "SELECT  PT." & strCampo & " intNumeroGuia " & _
             "FROM " & strTabela & " PT "
   
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF And Not IsNull(adoResultado("intNumeroGuia").Value) Then
                
            glngRetornaProximoNumeroGuia = adoResultado("intNumeroGuia").Value
            
        Else
            ExibeMensagem "Não foi possível retornar o número de Guia."
            Exit Function
        End If
    End If

    strSql = "UPDATE " & strTabela & " SET " & strCampo & " = " & strCampo & " + 1 WHERE " & strCampo & " = " & glngRetornaProximoNumeroGuia
    
    Set adoCommand = New ADODB.Command
    Set adoCommand.ActiveConnection = gcncADOMain
    adoCommand.CommandText = strSql
    adoCommand.Execute intRegistros
    
    If intRegistros = 0 Then GoTo ProximoNumeroGuia
    
    Exit Function

Problema_Na_Rotina:
    ExibeMensagem "Não foi possível concluir a operação de retorno de número de Guia."
    gobjBanco.ExecutaRollbackTrans
    
End Function

Public Sub LePlanoContaGeralBanco(cboCodigo As ComboBox, _
                             cboDescricao As ComboBox, _
                  ParamArray Parametro())
    Dim vntItem             As Variant
    Dim strAux              As String
    Dim strSql              As String
    Dim strCondicao         As String
    Dim blnConcatenaItem    As Boolean
    Dim blnSoSintetica      As Boolean
    Dim adoResultado        As ADODB.Recordset
    cboCodigo.Clear
    cboDescricao.Clear
    strAux = ""
    For Each vntItem In Parametro
        If InStr(UCase(vntItem), "SELECT") = 0 Then
            If Trim(strAux) = "" Then
                strAux = strAux & " WHERE "
            ElseIf blnConcatenaItem Then
                If Trim(UCase(strCondicao)) = "OU" Then
                    strAux = strAux & " OR "
                Else
                    strAux = strAux & " AND "
                End If
            End If
            If Trim(vntItem) = "OU" Or Trim(vntItem) = "ST" Then
                blnConcatenaItem = False
            Else
                blnConcatenaItem = True
            End If
        End If
        Select Case UCase(vntItem)
        Case "FN" 'Financeira
            strAux = strAux & "ABS(PC.blnFinanceira) = 1"
        Case "RT" 'Retenção
            strAux = strAux & "ABS(PC.blnRetencao) = 1"
        Case "EO" 'Extra-orçamentária
            strAux = strAux & "ABS(PC.blnExtraOrcamentaria) = 1"
        Case "IB" 'Integrar balanço
            strAux = strAux & "ABS(PC.blnIntegraBalanco) = 1"
        Case "RF" 'Retificadora
            strAux = strAux & "ABS(PC.blnRetificadora) = 1"
        Case "IS" 'Inversão de saldo
            strAux = strAux & "ABS(PC.blnInversaoDeSaldo) = 1"
        Case "CR" 'Conta de natureza credora
            strAux = strAux & "ABS(PC.blnNaturezaDaConta) = 0"
        Case "DC" 'Disponibilidade de Caixa
            strAux = strAux & "ABS(PC.bytdisponibilidadedecaixa) = 1"
        Case "DV" 'Conta de natureza devedora
            strAux = strAux & "ABS(PC.blnNaturezaDaConta) = 1"
        Case "OU"
            strCondicao = Trim(UCase(vntItem))
        Case "ST"
            blnSoSintetica = True
        Case Else
            strAux = strAux & vntItem
        End Select
    Next
    If blnSoSintetica = False Then
        If Trim(strAux) <> "" Then
            strAux = strAux & " AND blnAnalitica = 1 "
        Else
            strAux = strAux & "WHERE blnAnalitica = 1 "
        End If
    End If
    If InStr(UCase(strAux), "SELECT") <> 0 Then
        strSql = strAux
    Else
        strSql = ""
        strSql = strSql & "SELECT PC.PKId, CB.intNumeroConta, PC.strDescricao "
        strSql = strSql & "FROM "
        strSql = strSql & gstrPlanoConta & " PC, "
        strSql = strSql & gstrContaBancaria & " CB "
        strSql = strSql & strAux & " AND CB.PKId = PC.intContaBancaria ORDER BY CB.intNumeroConta"
    End If
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 60, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                cboDescricao.AddItem !strDescricao
                cboDescricao.ItemData(cboDescricao.NewIndex) = !Pkid
                cboCodigo.AddItem !intNumeroConta
                cboCodigo.ItemData(cboCodigo.NewIndex) = !Pkid
                .MoveNext
            Loop
        End With
    End If
End Sub


Public Sub leCodigoEvento(txtCampo As Object, ByVal cboCampo As Object)
    Dim strSql        As String
    Dim Pkid          As Integer
    Dim adoResultado  As New ADODB.Recordset
    
    Pkid = gstrItemData(cboCampo)
    
    If Pkid = 0 Then
        txtCampo.Text = ""
        Exit Sub
    End If
    
    strSql = ""
   
    strSql = strSql & "SELECT strCodigo FROM "
    strSql = strSql & gstrEvento
    strSql = strSql & " WHERE PKID = " & CStr(Pkid)
   
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSql, 60, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            txtCampo.Text = adoResultado!strCodigo
         End If
      End With
   End If
End Sub


Public Sub glIgualaContas(cboClicado As ComboBox, _
                          cboAIgualar As ComboBox, _
                 Optional lvw_Lista As ListView, _
                 Optional blnAlterando As Boolean)
    cboAIgualar.ListIndex = gintIndiceCBO(cboAIgualar, _
                                          gstrItemData(cboClicado))
    If lvw_Lista Is Nothing = False Then
        If gblnEncontroItemNoListView(lvw_Lista, gstrItemData(cboClicado), lvwTag) Then
            blnAlterando = True
        Else
            blnAlterando = False
        End If
    End If
End Sub


Public Sub LeSaldoConvenio(intConvenio As Integer, _
                           bytTipo As Byte, _
                           txtSaldo As TextBox, _
                  Optional txtData As TextBox, _
                  Optional txtValor As TextBox, _
                  Optional txtTotal As TextBox)
    '---------------------------------------------------------------
    ' SUB USADA PARA LÊ INFORMAÇÕES DO CONVÊNIO
    '---------------------------------------------------------------
    ' PARÂMETRO:
    '
    ' 1 intConvenio - Chave do convênio (PKId)
    ' 2 bytTipo - indica se é para empenho (despesa - bytTipo = 0)
    '             ou arrecadação (receita - bytTipo = 1)
    ' 3 txtSaldo - Saldo do empenho (para arrecadação ou empenho)
    ' 4 txtData - Data limite para aplicação do convênio
    ' 5 txtValor - Valor do Convêmio
    ' 6 txtTotal - Total arrecadado ou empenhado
    '---------------------------------------------------------------

    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    txtSaldo = ""
    If txtData Is Nothing = False Then
        txtData = ""
    End If
    If txtValor Is Nothing = False Then
        txtValor = ""
    End If
    If txtTotal Is Nothing = False Then
        txtTotal = ""
    End If
    strSql = ""
    
    strSql = strSql & gstrStoredProcedure("sp_LeSaldoConvenio", _
        intConvenio & ", " & bytTipo, True)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                txtSaldo = gstrConvVrDoSql(!dblSaldo)
                If txtData Is Nothing = False Then
                    txtData = gstrDataFormatada(!dtmDataFinal)
                End If
                If txtValor Is Nothing = False Then
                    txtValor = gstrConvVrDoSql(!dblValor)
                End If
                If txtTotal Is Nothing = False Then
                    txtTotal = gstrConvVrDoSql(!dblTotal)
                End If
            End If
        End With
    End If
End Sub

Public Sub LePrevisaoReceitaGeral(cboCodigo As ComboBox, _
                                  cboDescricao As ComboBox, _
                         Optional strQuery As String)
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    cboCodigo.Clear
    cboDescricao.Clear
    If Trim(strQuery) = "" Then
        strSql = ""
        strSql = strSql & "SELECT CO.PKId, CO.strCodigoOrcamentario, CO.strDescricao "
        strSql = strSql & "FROM "
        strSql = strSql & gstrCodigoOrcamentario & " CO, "
        strSql = strSql & gstrPrevisaoDaReceita & " PR "
        strSql = strSql & "WHERE CO.PKId = PR.intCodigoOrcamentario "
        strSql = strSql & "AND PR.intExercicio = " & gintExercicio & " "
        strSql = strSql & "ORDER BY CO.strDescricao"
    Else
        strSql = strQuery
    End If
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                cboDescricao.AddItem !strDescricao
                cboDescricao.ItemData(cboDescricao.NewIndex) = !Pkid
                cboCodigo.AddItem gvntFormatacaoEspecifica(!strCodigoOrcamentario)
                cboCodigo.ItemData(cboCodigo.NewIndex) = !Pkid
                .MoveNext
            Loop
        End With
    End If
End Sub


Public Sub LePlanoContaGeral(cboCodigo As ComboBox, _
                             cboDescricao As ComboBox, _
                  ParamArray Parametro())
    Dim vntItem             As Variant
    Dim strAux              As String
    Dim strSql              As String
    Dim strCondicao         As String
    Dim blnConcatenaItem    As Boolean
    Dim blnSoSintetica      As Boolean
    Dim adoResultado        As ADODB.Recordset
    cboCodigo.Clear
    cboDescricao.Clear
    strAux = ""
    For Each vntItem In Parametro
        If InStr(UCase(vntItem), "SELECT") = 0 Then
            If Trim(strAux) = "" Then
                strAux = strAux & " WHERE "
            ElseIf blnConcatenaItem Then
                If Trim(UCase(strCondicao)) = "OU" Then
                    strAux = strAux & " OR "
                Else
                    strAux = strAux & " AND "
                End If
            End If
            If Trim(vntItem) = "OU" Or Trim(vntItem) = "ST" Then
                blnConcatenaItem = False
            Else
                blnConcatenaItem = True
            End If
        End If
        Select Case UCase(vntItem)
        Case "FN" 'Financeira
            strAux = strAux & "ABS(PC.blnFinanceira) = 1"
        Case "RT" 'Retenção
            strAux = strAux & "ABS(PC.blnRetencao) = 1"
        Case "EO" 'Extra-orçamentária
            strAux = strAux & "ABS(PC.blnExtraOrcamentaria) = 1"
        Case "IB" 'Integrar balanço
            strAux = strAux & "ABS(PC.blnIntegraBalanco) = 1"
        Case "RF" 'Retificadora
            strAux = strAux & "ABS(PC.blnRetificadora) = 1"
        Case "IS" 'Inversão de saldo
            strAux = strAux & "ABS(PC.blnInversaoDeSaldo) = 1"
        Case "CR" 'Conta de natureza credora
            strAux = strAux & "ABS(PC.blnNaturezaDaConta) = 0"
        Case "DV" 'Conta de natureza devedora
            strAux = strAux & "ABS(PC.blnNaturezaDaConta) = 1"
        Case "PA" 'Contas Patrimoniais
            strAux = strAux & "ABS(PC.Blnpatrimonial) = 1"
        Case "OU"
            strCondicao = Trim(UCase(vntItem))
        Case "ST"
            blnSoSintetica = True
        Case Else
            strAux = strAux & vntItem
        End Select
    Next
    If blnSoSintetica = False Then
        If Trim(strAux) <> "" Then
            strAux = strAux & " AND blnAnalitica = 1 "
        Else
            strAux = strAux & "WHERE blnAnalitica = 1 "
        End If
    End If
    If InStr(UCase(strAux), "SELECT") <> 0 Then
        strSql = strAux
    Else
        strSql = ""
        strSql = strSql & "SELECT PC.PKId, PC.strContaContabil, PC.strDescricao "
        strSql = strSql & "FROM "
        strSql = strSql & gstrPlanoConta & " PC "
        strSql = strSql & strAux & " ORDER BY PC.strDescricao"
    End If
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 60, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                cboDescricao.AddItem !strDescricao
                cboDescricao.ItemData(cboDescricao.NewIndex) = !Pkid
                cboCodigo.AddItem gvntFormatacaoEspecifica(!strContaContabil)
                cboCodigo.ItemData(cboCodigo.NewIndex) = !Pkid
                .MoveNext
            Loop
        End With
    End If
End Sub

Public Sub PreencheEventobyCodigo(txtCampo As Object, ByVal cboCampo As Object, strTipoEvento As String)
    Dim strSql        As String
    Dim Pkid          As Integer
    Dim adoResultado  As New ADODB.Recordset
    Dim strMantemValor As String
    
    If Trim(txtCampo.Text) = "" Then
        cboCampo.ListIndex = -1
        Exit Sub
    End If
    
    If cboCampo.ListCount = 0 Then
        LeDaTabelaParaObj gstrEvento, cboCampo, "SELECT PKID, strDescricao FROM " & gstrEvento & _
        IIf(strTipoEvento <> "todos", " WHERE intTipoEvento in (" & strTipoEvento & ")", "")
    End If
    
    strSql = ""
   
    strSql = strSql & "SELECT PKID FROM "
    strSql = strSql & gstrEvento
    strSql = strSql & " WHERE strCodigo = " & Trim(txtCampo.Text)
    If strTipoEvento <> "todos" Then
        strSql = strSql & " AND intTipoEvento in (" & strTipoEvento & ")"
    End If
   
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSql, 60, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            cboCampo.ListIndex = gintIndiceCBO(cboCampo, adoResultado!Pkid)
            txtCampo.SelStart = 0
            txtCampo.SelLength = Len(txtCampo.Text)
         Else
            strMantemValor = txtCampo.Text
            cboCampo.ListIndex = -1
            txtCampo.Text = strMantemValor
            txtCampo.SelStart = 0
            txtCampo.SelLength = Len(txtCampo.Text)
         End If
      End With
   End If
End Sub


Public Function gstrDigitoReceita() As String
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = "SELECT strDigitoReceita  FROM " & gstrConfiguracaoGeral
    strSql = strSql & " WHERE PKID = ("
    strSql = strSql & "SELECT MAX(PKID) PKID FROM " & gstrConfiguracaoGeral & ")"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not adoResultado.EOF Then
                If IsNull(!strDigitoReceita) Or !strDigitoReceita = "" Then
                    gstrDigitoReceita = "0"
                Else
                    gstrDigitoReceita = adoResultado!strDigitoReceita
                End If
            End If
        End With
    End If
    
    adoResultado.Close: Set adoResultado = Nothing
   
End Function

Public Function gstrDigitoDespesa() As String
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = "SELECT strDigitoDespesa  FROM " & gstrConfiguracaoGeral
    strSql = strSql & " WHERE PKID = ("
    strSql = strSql & "SELECT MAX(PKID) PKID FROM " & gstrConfiguracaoGeral & ")"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not adoResultado.EOF Then
                If IsNull(!strDigitoDespesa) Or !strDigitoDespesa = "" Then
                    gstrDigitoDespesa = "0"
                Else
                    gstrDigitoDespesa = adoResultado!strDigitoDespesa
                End If
            End If
        End With
    End If
    
    adoResultado.Close: Set adoResultado = Nothing
   
End Function

Public Function BuscaCodigosPeloEvento(intEvento As Integer, strPrimeiroDigito As String, strTpMovimento As String, intTpEvento As Byte) As String
   
   Dim strSql      As String
   Dim adoResultado As New ADODB.Recordset
   Dim intContador  As Integer
   
   strSql = "SELECT EV.PKID, EV.strDescricao, EVC.intContaContabil, PC.strContaContabil"
   strSql = strSql & " FROM " & gstrEvento & " EV, " & IIf(strTpMovimento = "C", gstrEventoContaContabilCredito, gstrEventoContaContabilDebito) & " EVC,"
   strSql = strSql & gstrPlanoConta & " PC"
   strSql = strSql & " WHERE  EV.Pkid = EVC.intEvento AND EV.PKID = " & intEvento & " AND "
   strSql = strSql & " EVC.intContaContabil = PC.PKid AND "
   strSql = strSql & strSUBSTRING & "(PC.strContaContabil,1," & Len(Trim(strPrimeiroDigito)) & ") = '" & strPrimeiroDigito & "' AND "
   strSql = strSql & " EV.intTipoEvento = " & intTpEvento
   
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSql, 15, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            BuscaCodigosPeloEvento = Mid(!strContaContabil, Len(Trim(strPrimeiroDigito)) + 1)
            For intContador = 1 To Len(Mid(adoResultado!strContaContabil, Len(Trim(strPrimeiroDigito)) + 1))
               'Vamos sair do loop caso encontre valor <> 0 e o len do resultado < que 6 por causa dos registros criados pelo Flavio (Ex: 33000000001)
               If Mid(BuscaCodigosPeloEvento, Len(BuscaCodigosPeloEvento), 1) = "0" Or Len(BuscaCodigosPeloEvento) > 6 Then
                  BuscaCodigosPeloEvento = Mid(BuscaCodigosPeloEvento, 1, Len(BuscaCodigosPeloEvento) - 1)
               Else
                  Exit For
               End If
            Next
         End If
      End With
   End If
   
End Function


Public Sub LeTabelaContaDoEvento(gstrTabela As String, _
                                 lvw_Lista As ListView, _
                                 intEvento As Integer)
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    Dim objLista        As Object
    lvw_Lista.ListItems.Clear
    strSql = ""
    strSql = strSql & "SELECT PC.PKId, PC.strContaContabil, PC.strDescricao , EC.bytContaGrupo "
    strSql = strSql & "FROM " & gstrPlanoConta & " PC, " & gstrTabela & " EC "
    strSql = strSql & "WHERE PC.PKId = EC.intContaContabil "
    strSql = strSql & "AND EC.intEvento = " & intEvento & " "
    strSql = strSql & "ORDER BY PC.strDescricao"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                Set objLista = lvw_Lista.ListItems.Add(, , gvntFormatacaoEspecifica(!strContaContabil))
                objLista.SubItems(1) = Trim(!strDescricao)
                objLista.SubItems(2) = IIf(!bytContaGrupo = 0, "Não", "Sim")
                objLista.Tag = !Pkid
                .MoveNext
            Loop
            .Close
        End With
    End If
End Sub

Public Function gstrMontaCodigoBarras(bytTipoGuia As Byte, intContaBancaria As Long, dblValorGuia, dtmDataVencimento As String, intFebraban As Integer, INTNUMERO As Long, blnValorVariavel As Boolean, blnValorEmReal As Boolean)
Dim adoResultado As New ADODB.Recordset
Dim strCodBarras As String
Dim dblValorMult As Double
Dim bytDigito    As Byte
    
    'Vamos verificar quantas casas decimais possui o valor para aplicar a multiplicacao correta no codigo de barras
    dblValorMult = Val("1" + String(Len(Trim(dblValorGuia)) - InStr(1, dblValorGuia, ","), "0"))
    
    'Tipo Febraban
    If bytTipoGuia = FEBRABAN Then
        strCodBarras = ""
        strCodBarras = "81" 'Digito fixo
        strCodBarras = strCodBarras & IIf(blnValorVariavel, "7", "6")
        strCodBarras = strCodBarras & Format$((dblValorGuia * dblValorMult), "00000000000")   'Valor da guia
        strCodBarras = strCodBarras & Format$(intFebraban, "0000") 'Codigo do Febraban
        strCodBarras = strCodBarras & Replace(Format$(dtmDataVencimento, "YYYY/MM/DD"), "/", "") 'Vencimento da Guia
        strCodBarras = strCodBarras & "0000" 'Conta bancaria tipo nulo que é a do Febrabraban
        strCodBarras = strCodBarras & Format$(INTNUMERO, "000000000") 'Número sequencial da guia
        strCodBarras = strCodBarras & Year(gstrDataDoSistema) 'Exercício corrente
        
        bytDigito = gstrCalculaDigitoModulo10(strCodBarras) 'Calcula o digito
        strCodBarras = Mid(strCodBarras, 1, 3) & bytDigito & Mid(strCodBarras, 4, Len(strCodBarras)) 'Adiciona o digito ao codigo de barras
        
    'Tipo Ficha Compensação
    Else
        'Vamos obter o tipo de Codigo de Barras
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO("SELECT CB.strCedente, CB.strConta, CB.intTipoCodigoBarra, CB.strDigitoVerificador, AG.strAgencia, TC.strEspecieDoc, TC.strCarteira FROM " & gstrContaBancaria & " CB, " & gstrAgencia & " AG, " & gstrTipoCodigoBarra & " TC WHERE CB.intAgencia = AG.Pkid and TC.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " CB.intTipoCodigoBarra and CB.Pkid = " & intContaBancaria, 5, adoResultado) Then
            
            Select Case adoResultado("intTipoCodigoBarra").Value
            
                Case Is = 1 'Banespa
                    
                    'Vamos verificar se a conta 7 posicoes
                    If Len(Replace(Replace(Replace(UCase(adoResultado("strConta").Value), "-", ""), ".", ""), "X", "")) <> 7 Then
                        strCodBarras = ""
                        ExibeMensagem "O Número da Conta deve ter 7 dígitos. Conta: " & adoResultado("strConta").Value
                        Exit Function
                    End If
                    
                    'Vamos verificar se existe digito
                    If Len(Trim(adoResultado("strDigitoVerificador").Value)) = 0 Or Not IsNumeric(adoResultado("strDigitoVerificador").Value) Then
                        strCodBarras = ""
                        ExibeMensagem "O Dígito da Agência deve ser informado numericamente. Conta: " & adoResultado("strConta").Value
                        Exit Function
                    End If
                    
                    'Vamos verificar se a agencia tem mais de 3 posicoes
                    If Len(Replace(Replace(Replace(UCase(adoResultado("strAgencia").Value), "-", ""), ".", ""), "X", "")) > 3 Then
                        strCodBarras = ""
                        ExibeMensagem "O Número da Agência deve ter no máximo 3 dígitos. Conta: " & adoResultado("strConta").Value & " Agência: " & adoResultado("strAgencia").Value
                        Exit Function
                    End If
                    
                    strCodBarras = ""
                    strCodBarras = "033" 'Banco
                    strCodBarras = strCodBarras & IIf(blnValorEmReal, "9", "8")  'Cod Moeda - 9 Real, 8 Outras moedas
                    strCodBarras = strCodBarras & Format(DateDiff("d", "07/10/1997", dtmDataVencimento), "0000") 'Fator Vencimento
                    strCodBarras = strCodBarras & Format$((dblValorGuia * dblValorMult), "0000000000") 'Valor da guia
                    strCodBarras = strCodBarras & Format$(Replace(Replace(Replace(UCase(adoResultado("strAgencia").Value), "-", ""), ".", ""), "X", ""), "000")   'Codigo da Agencia
                    strCodBarras = strCodBarras & Format$(Replace(Replace(Replace(UCase(adoResultado("strConta").Value), "-", ""), ".", ""), "X", ""), "0000000")   'Conta bancaria
                    strCodBarras = strCodBarras & Format$(adoResultado("strDigitoVerificador").Value, "0")  'Digito
                    strCodBarras = strCodBarras & Format$(INTNUMERO, "0000000") 'Nosso número sequencial da guia
                    strCodBarras = strCodBarras & "00" 'Filler
                    strCodBarras = strCodBarras & "033" 'Banco
                    strCodBarras = strCodBarras & gstrCalculaDigitoModulo10(Mid(strCodBarras, 19, 25)) 'Digito verificador 1
                    strCodBarras = Mid(strCodBarras, 1, Len(strCodBarras) - 1) & gstrCalculaDigito2Asbace(Mid(strCodBarras, 19, 25)) 'Digito verificador 2
                    'strCodBarras = Mid(strCodBarras, 1, Len(strCodBarras) - 1) & gstrCalculaDigito2Asbace(strCodBarras) 'Digito verificador 2
                    strCodBarras = Mid(strCodBarras, 1, 4) & gstrCalculaDigitoAutoConferencia(strCodBarras) & Mid(strCodBarras, 5, 40) 'Digito de autoconferencia
                
                Case Is = 2 'Caixa Economica Federal - Sem Registro com 16 Posições
                
                    'Vamos verificar se a conta tem mais de 5 posicoes
                    If Len(Replace(Replace(Replace(UCase(adoResultado("strConta").Value), "-", ""), ".", ""), "X", "")) > 5 Then
                        strCodBarras = ""
                        ExibeMensagem "O Número da Conta deve ter no máximo 5 dígitos. Conta: " & adoResultado("strConta").Value
                        Exit Function
                    End If
                    
                    strCodBarras = ""
                    strCodBarras = "104" 'Banco
                    strCodBarras = strCodBarras & IIf(blnValorEmReal, "9", "8")  'Cod Moeda - 9 Real, 8 Outras moedas
                    strCodBarras = strCodBarras & Format(DateDiff("d", "07/10/1997", dtmDataVencimento), "0000") 'Fator Vencimento
                    strCodBarras = strCodBarras & Format$((dblValorGuia * dblValorMult), "0000000000") 'Valor da guia
                    strCodBarras = strCodBarras & Format$(Replace(Replace(Replace(UCase(adoResultado("strConta").Value), "-", ""), ".", ""), "X", ""), "00000")     'Conta bancaria
                    strCodBarras = strCodBarras & Format$(Replace(Replace(Replace(UCase(adoResultado("strAgencia").Value), "-", ""), ".", ""), "X", ""), "0000")    'Codigo da Agencia
                    strCodBarras = strCodBarras & "8" 'Codigo da carteira
                    strCodBarras = strCodBarras & "7" 'Constante
                    strCodBarras = strCodBarras & Format$(INTNUMERO, "00000000000000") 'Nosso número sequencial da guia
                    strCodBarras = Mid(strCodBarras, 1, 4) & gstrCalculaDigitoAutoConferencia(strCodBarras) & Mid(strCodBarras, 5, 40) 'Digito de autoconferencia
                
                Case Is = 3 'Caixa Economica Federal - Sem Registro
                
                    'Vamos verificar se o cedente tem 12 posicoes
                    If Len(Replace(Replace(Replace(UCase(Space$(0) & adoResultado("strCedente").Value), "-", ""), ".", ""), "X", "")) <> 12 Then
                        strCodBarras = ""
                        ExibeMensagem "O Cedente deve ter 12 dígitos. Conta: " & adoResultado("strConta").Value
                        Exit Function
                    End If
                    
                    strCodBarras = ""
                    strCodBarras = "104" 'Banco
                    strCodBarras = strCodBarras & IIf(blnValorEmReal, "9", "8")  'Cod Moeda - 9 Real, 8 Outras moedas
                    strCodBarras = strCodBarras & Format(DateDiff("d", "07/10/1997", dtmDataVencimento), "0000") 'Fator Vencimento
                    strCodBarras = strCodBarras & Format$((dblValorGuia * dblValorMult), "0000000000") 'Valor da guia
                    strCodBarras = strCodBarras & "82"
                    strCodBarras = strCodBarras & Format$(INTNUMERO, "00000000")  'Nosso número sequencial da guia
                    strCodBarras = strCodBarras & Format$(Replace(Replace(Replace(UCase(adoResultado("strAgencia").Value), "-", ""), ".", ""), "X", ""), "0000") 'Codigo da Agencia
                    strCodBarras = strCodBarras & Format$(Replace(Replace(Replace(UCase(Mid(adoResultado("strCedente").Value, 1, 11)), "-", ""), ".", ""), "X", ""), "00000000000") 'Cedente
                    strCodBarras = Mid(strCodBarras, 1, 4) & gstrCalculaDigitoAutoConferencia(strCodBarras) & Mid(strCodBarras, 5, 40) 'Digito de autoconferencia
                
                Case Is = 4 'Banco do Brasil - Vinculados a convenios com numeracao superior a 1.000.000
                
                    strCodBarras = ""
                    strCodBarras = "001" 'Banco
                    strCodBarras = strCodBarras & "9"   'Cod Moeda - 9 Real
                    strCodBarras = strCodBarras & Format(DateDiff("d", "07/10/1997", dtmDataVencimento), "0000") 'Fator Vencimento
                    strCodBarras = strCodBarras & Format$((dblValorGuia * dblValorMult), "0000000000") 'Valor da guia
                    strCodBarras = strCodBarras & "000000" 'Zeros
                    strCodBarras = strCodBarras & Format$(adoResultado("strEspecieDoc").Value, "0000000") 'Numero convenio
                    strCodBarras = strCodBarras & Format$(INTNUMERO, "0000000000")  'Nosso número
                    strCodBarras = strCodBarras & Format$(adoResultado("strCarteira").Value, "00") 'Carteira
                    strCodBarras = Mid(strCodBarras, 1, 4) & gstrCalculaDigitoAutoConferencia(strCodBarras) & Mid(strCodBarras, 5, 40) 'Digito de autoconferencia
                
                Case Else   'Não relacionado
                    strCodBarras = ""
                    ExibeMensagem "Não foi relacionado nenhum tipo de código de barras para a Conta Bancária " & adoResultado("strConta").Value
                    
            End Select
            
        End If
    
    End If
    
    gstrMontaCodigoBarras = strCodBarras
    
End Function

Public Function gstrMontaLinhaDigitavel(bytTipoGuia As Byte, strCodigoBarras As String)
Dim strNumeroBoleto1     As String
Dim strNumeroBoleto2     As String
Dim strNumeroBoleto3     As String
Dim strNumeroBoleto4     As String
Dim strNumeroBoleto5     As String
Dim strLinhaDigitavel    As String

    If bytTipoGuia = FEBRABAN Then
    
        strNumeroBoleto1 = Mid(strCodigoBarras, 1, 11) & "-" & gstrCalculaDigitoModulo10(Mid(strCodigoBarras, 1, 11))
        strNumeroBoleto2 = Mid(strCodigoBarras, 12, 11) & "-" & gstrCalculaDigitoModulo10(Mid(strCodigoBarras, 12, 11))
        strNumeroBoleto3 = Mid(strCodigoBarras, 23, 11) & "-" & gstrCalculaDigitoModulo10(Mid(strCodigoBarras, 23, 11))
        strNumeroBoleto4 = Mid(strCodigoBarras, 34, 11) & "-" & gstrCalculaDigitoModulo10(Mid(strCodigoBarras, 34, 11))
        strLinhaDigitavel = strNumeroBoleto1 & " " & strNumeroBoleto2 & " " & strNumeroBoleto3 & " " & strNumeroBoleto4
    
    Else
        
        strNumeroBoleto1 = Mid(strCodigoBarras, 1, 4) & Mid(strCodigoBarras, 20, 5) & gstrCalculaDigitoModulo10(Mid(strCodigoBarras, 1, 4) & Mid(strCodigoBarras, 20, 5))
        strNumeroBoleto2 = Mid(strCodigoBarras, 25, 10) & gstrCalculaDigitoModulo10(Mid(strCodigoBarras, 25, 10))
        strNumeroBoleto3 = Mid(strCodigoBarras, 35, 10) & gstrCalculaDigitoModulo10(Mid(strCodigoBarras, 35, 10))
        strNumeroBoleto4 = Mid(strCodigoBarras, 5, 1)
        strNumeroBoleto5 = Mid(strCodigoBarras, 6, 14)
        strLinhaDigitavel = Mid(strNumeroBoleto1, 1, 5) & "." & Mid(strNumeroBoleto1, 6, 5) & " " & Mid(strNumeroBoleto2, 1, 5) & "." & Mid(strNumeroBoleto2, 6, 6) & " " & Mid(strNumeroBoleto3, 1, 5) & "." & Mid(strNumeroBoleto3, 6, 6) & " " & strNumeroBoleto4 & " " & strNumeroBoleto5
    
    End If
    
    gstrMontaLinhaDigitavel = strLinhaDigitavel
    
End Function

Public Function gstrMontaNossoNumero(intContaBancaria As Long, INTNUMERO As Long)
Dim adoResultado   As New ADODB.Recordset
Dim strNossoNumero As String
    
    'Vamos obter o tipo de Codigo de Barras
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO("SELECT CB.strConta, CB.strDigitoVerificador, CB.intTipoCodigoBarra, AG.strAgencia, TC.strEspecieDoc FROM " & gstrContaBancaria & " CB, " & gstrAgencia & " AG, " & gstrTipoCodigoBarra & " TC WHERE CB.intAgencia = AG.Pkid and TC.Pkid = CB.intTipoCodigoBarra and CB.Pkid = " & intContaBancaria, 5, adoResultado) Then
        
        Select Case adoResultado("intTipoCodigoBarra").Value
        
            Case Is = 1 'Banespa
                
                'Vamos verificar se a conta tem mais de 7 posicoes
                If Len(Replace(Replace(Replace(UCase(adoResultado("strConta").Value), "-", ""), ".", ""), "X", "")) > 7 Then
                    strNossoNumero = ""
                    ExibeMensagem "O Número da Conta deve ter no máximo 7 dígitos. Conta: " & adoResultado("strConta").Value
                    Exit Function
                End If
                
                'Vamos verificar se o digito da agencia foi informado
                If Len(Trim(adoResultado("strDigitoVerificador").Value)) = 0 Or Not IsNumeric(adoResultado("strDigitoVerificador").Value) Then
                    strNossoNumero = ""
                    ExibeMensagem "O Dígito da Agência deve ser informado numericamente. Conta: " & adoResultado("strConta").Value
                    Exit Function
                End If
                
                'Vamos verificar se a agencia tem mais de 3 posicoes
                If Len(Replace(Replace(Replace(UCase(adoResultado("strAgencia").Value), "-", ""), ".", ""), "X", "")) > 3 Then
                    strNossoNumero = ""
                    ExibeMensagem "O Número da Agência deve ter no máximo 3 dígitos. Conta: " & adoResultado("strConta").Value & " Agência: " & adoResultado("strAgencia").Value
                    Exit Function
                End If
                
                strNossoNumero = ""
                strNossoNumero = strNossoNumero & Format$(Replace(Replace(Replace(UCase(adoResultado("strAgencia").Value), "-", ""), ".", ""), "X", ""), "000")   'Codigo da Agencia
                strNossoNumero = strNossoNumero & Format$(INTNUMERO, "0000000")   'Nosso número sequencial da guia
                strNossoNumero = strNossoNumero & gstrCalculaDigitoNossoNumero(strNossoNumero)   'Digito verificador
                
                strNossoNumero = Mid(strNossoNumero, 1, 3) & " " & Mid(strNossoNumero, 4, 7) & " " & Right(strNossoNumero, 1)
                
            Case Is = 2 'Caixa Economica Federal - Sem Registro com 16 Posições
            
                'Vamos verificar se a conta tem mais de 5 posicoes
                If Len(Replace(Replace(Replace(UCase(adoResultado("strConta").Value), "-", ""), ".", ""), "X", "")) > 5 Then
                    strNossoNumero = ""
                    ExibeMensagem "O Número da Conta deve ter no máximo 5 dígitos. Conta: " & adoResultado("strConta").Value
                    Exit Function
                End If
                
                strNossoNumero = ""
                strNossoNumero = strNossoNumero & "8" & Format$(INTNUMERO, "00000000000000")   'Nosso número sequencial da guia
                strNossoNumero = strNossoNumero & gstrCalculaDigitoAutoConferencia(strNossoNumero)   'Digito verificador
                
            Case Is = 3 'Caixa Economica Federal - Sem Registro
            
                'Vamos verificar se a conta tem mais de 5 posicoes
                If Len(Replace(Replace(Replace(UCase(adoResultado("strConta").Value), "-", ""), ".", ""), "X", "")) > 8 Then
                    strNossoNumero = ""
                    ExibeMensagem "O Número da Conta deve ter no máximo 8 dígitos. Conta: " & adoResultado("strConta").Value
                    Exit Function
                End If
                
                strNossoNumero = ""
                strNossoNumero = strNossoNumero & "82" & Format$(INTNUMERO, "00000000")   'Nosso número sequencial da guia
                strNossoNumero = strNossoNumero & gstrCalculaDigitoAutoConferencia(strNossoNumero)   'Digito verificador
                
            Case Is = 4 'Banco do Brasil - Vinculados a convenios com numeracao superior a 1.000.000
            
                strNossoNumero = ""
                strNossoNumero = strNossoNumero & Format$(adoResultado("strEspecieDoc").Value, "0000000")    'Numero Convenio
                strNossoNumero = strNossoNumero & Format$(INTNUMERO, "0000000000")   'Nosso número sequencial da guia
                
                
            Case Else   'Não relacionado
                strNossoNumero = ""
                ExibeMensagem "Não foi relacionado nenhum tipo de código de barras para a Conta Bancária " & adoResultado("strConta").Value
                
        End Select
        
    End If

    gstrMontaNossoNumero = strNossoNumero
    
End Function

Public Function gstrRetornaOps(strPagamentoPKID As String) As String
    Dim strSql As String
    Dim adoResultado  As ADODB.Recordset
    Dim i As Integer
    
    
    strSql = ""
    strSql = strSql & " SELECT DISTINCT op.intnumero "
    strSql = strSql & " FROM " & gstrPagamentoEstornoEmpenho & " PEE "
    strSql = strSql & " , " & gstrOrdemPagamento & " OP "
    strSql = strSql & " WHERE "
    strSql = strSql & " OP.PKID = PEE.INTORDEMPAGAMENTO AND "
    strSql = strSql & " PEE.INTPROCESSO = " & strPagamentoPKID
    
    strSql = strSql & " UNION SELECT DISTINCT op.intnumero "
    strSql = strSql & " FROM " & gstrOrdemPagamentoAnulacaoReceita & " PEE "
    strSql = strSql & " , " & gstrOrdemPagamento & " OP "
    strSql = strSql & " WHERE "
    strSql = strSql & " OP.PKID = PEE.INTORDEMPAGAMENTO AND "
    strSql = strSql & " PEE.INTPROCESSO = " & strPagamentoPKID

    
    
    
    Set gobjBanco = New clsBanco
   
    If gobjBanco.CriaADO(strSql, 60, adoResultado) Then
        While Not adoResultado.EOF
            If Len(Trim(adoResultado!INTNUMERO)) > 0 Then
               gstrRetornaOps = gstrRetornaOps & Format(adoResultado!INTNUMERO, "000") & ", "
            End If
            
            adoResultado.MoveNext
        Wend
    End If
    
    If Trim(gstrRetornaOps) <> "" Then
        gstrRetornaOps = Trim(Mid(Trim(gstrRetornaOps), 1, Len(gstrRetornaOps) - 2))
        For i = Len(gstrRetornaOps) To 1 Step -1
            If Mid(gstrRetornaOps, i, 1) = "," Then
                gstrRetornaOps = Mid(gstrRetornaOps, 1, i - 1) & " e" & Mid(gstrRetornaOps, i + 1)
                Exit For
            End If
        Next
    End If
    
End Function

Public Function gblnValidaNumeroProtocolo() As Boolean
Dim strSql As String
Dim adoResultado As ADODB.Recordset

    strSql = ""
    strSql = "SELECT * FROM " & gstrParametroProtocolo
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            gblnValidaNumeroProtocolo = IIf(IsNull(adoResultado("blnValidarNumeroProtocolo")), False, adoResultado("blnValidarNumeroProtocolo"))
        End If
    End If
    
End Function

Public Function SalvarOperacaoUsuario(strTabela As String, strModoOperacao As String, frmForm As Form) As Boolean
'Parâmetros:
    'strTabela       => A tabela que recebera manutenção
    'strModoOperacao => (I)nclusão, (A)lteração, (E)xclusão
    'frmForm         => O formulário que trata a manutenção na tabela

Dim blnErro As Boolean
Dim adoResultado            As ADODB.Recordset
Dim strSql                  As String
Dim strLabel                As String
Dim strFinalString          As String
'Auxiliares para o loop no formulário
Dim i As Integer
Dim j As Integer

'Possui 4 colunas
    '0 => O nome do campo da tabela (strTabela)
    '1 => O valor do campo do formulário ("NULL" quando vazio)
    '2 => O título (Label) do campo no formulário
    '3 => O prefixo do objeto ("txt","chk","cbo",etc...)
    
'Tag de objetos com nº 1 => TextBox que utilizam máscara específica e guardam na tabela o valor
                            'Sem a máscara
Dim vetCampos() As String

'Auxiliar para chamar a DLL e tratar o retorno da mesma
Dim strAux As String

'A DLL que fará o tratamento das manutenções no banco de dados
Dim objFuncao As clsFuncoes

'Atribui o campo descrição do objeto que identifica-se no form
Dim strDescricao As String

On Error Resume Next
    
    With frmForm
        For i = 0 To .Controls.Count - 1
            If UCase(.Controls(i).Name) = "TXTSTRDESCRICAO" Or UCase(.Controls(i).Name) = "TXTSTRNOME" Then
                strDescricao = .Controls(i).Text
                Exit For
            End If
        Next
    End With

        
        'Cria uma nova instância da DLL
        Set objFuncao = New clsFuncoes
        
        'Primeiro índice do vetor (linha)
        j = 0
        
        With frmForm
            'Exclusão
            If strModoOperacao = "E" Then
                ReDim vetCampos(1, 0)
                vetCampos(0, 0) = "PKID"
                vetCampos(1, 0) = Val(.Controls("txtpkid").Text)
                strDescricao = .Controls("txtstrDescricao").Text
            Else
                'Percorre os controles do form
                For i = 0 To .Controls.Count - 1
                    
                    'Elimina os objetos que não tem relacionamento com os campos da tabela
                    'Estes controles são identificados por ter seu prefixo (tres letras) separados do
                    'nome por "_"
                    If Mid(.Controls(i).Name, 4, 1) <> "_" And (Not TypeOf .Controls(i) Is Label) Then
                        
                        If Not (TypeOf .Controls(i) Is OptionButton) Or .Controls(i) = True Then
                        
                            'If UCase(.Controls(i).Name) = "TXTSTRDESCRICAO" Or UCase(.Controls(i).Name) = "TXTSTRNOME" Then
                            '    strDescricao = .Controls(i).Text
                            'End If
                            
                            'Redimensiona o Vetor para o novo campo encontrado
                            ReDim Preserve vetCampos(3, j)
                            
                            'Elimina o prefixo do objeto do formulário para se descobrir o nome do campo na tabela
                            'ex:
                                'Nome do Campo no Form  = Nome do Campo na Tabela
                                'txtStrDescricao        = StrDescricao
                            strAux = Right(.Controls(i).Name, Len(.Controls(i).Name) - 3)
                            
                            'Atribui ao vetor o nome do campo na tabela
                            vetCampos(0, j) = strAux
                            
                            'Verifica se controle está preenchido
                            If Trim(.Controls(i)) = "" Then
                                vetCampos(1, j) = "NULL"
                            Else
                                If TypeOf .Controls(i) Is DTPicker Then
        '                            vetCampos(1, j) = Format(.Controls(i), "YYYY/MM/DD")
                                    If bytDBType = EDatabases.SQLServer Then
                                        vetCampos(1, j) = Format(.Controls(i), "YYYY/MM/DD")
                                    ElseIf bytDBType = EDatabases.Oracle Then
                                        vetCampos(1, j) = gstrFormataDataOracle(.Controls(i), "yyyy/mm/dd")
                                    End If
                                ElseIf TypeOf .Controls(i) Is MaskEdBox Then
                                    If UCase(Left(strAux, 3)) = "DTM" Then
                                        If Trim(.Controls(i).FormattedText) <> "/  /" Then
                                            If gblnDataValida(.Controls(i).FormattedText, True) Then
                                                'if len(.Controls(i).FormattedText)
        '                                        vetCampos(1, j) = Format(.Controls(i).FormattedText, "YYYY/MM/DD hh:mm:ss")
                                                If bytDBType = EDatabases.SQLServer Then
                                                    vetCampos(1, j) = Format(.Controls(i).FormattedText, "YYYY/MM/DD hh:mm:ss")
                                                ElseIf bytDBType = EDatabases.Oracle Then
                                                    vetCampos(1, j) = gstrFormataDataOracle(.Controls(i).FormattedText)
                                                End If
                                            Else
                                                'EXIT FUNCTION
                                            End If
                                        Else
                                            vetCampos(1, j) = "NULL"
                                        End If
                                    ElseIf InStr(1, strAux, "CNPJCPF") > 0 Then
                                        If Len(Trim(gstrValorSemMascara(gstrCGCCPFFormatado(.Controls(i).ClipText)))) = 11 Then
                                            If gblnCPFOk(.Controls(i).ClipText) Then
                                                vetCampos(1, j) = .Controls(i).ClipText
                                            Else
                                                If .Controls(i).Enabled Then
                                                    .Controls(i).SetFocus
                                                End If
                                                'EXIT FUNCTION
                                            End If
                                        ElseIf Len(Trim(gstrValorSemMascara(.Controls(i).ClipText))) = 14 Then
                                            If gblnCGCOk(.Controls(i).ClipText) Then
                                                vetCampos(1, j) = .Controls(i).ClipText
                                            Else
                                                If .Controls(i).Enabled Then
                                                    .Controls(i).SetFocus
                                                End If
                                                'EXIT FUNCTION
                                            End If
                                        Else
                                            If .Controls(i).Enabled Then
                                                .Controls(i).SetFocus
                                            End If
                                            'EXIT FUNCTION
                                        End If
                                    Else
                                        vetCampos(1, j) = .Controls(i).ClipText
                                    End If
                                ElseIf TypeOf .Controls(i) Is OptionButton Then
                                    vetCampos(1, j) = .Controls(i).Index
                                ElseIf TypeOf .Controls(i) Is DataCombo Then
                                    'Valor selecionado/digitado no DataCombo encontra-se na lista
                                    If .Controls(i).MatchedWithList Then
                                        'Guarda o código (BoundText) e o valor preenchido (Text) do DataCombo
                                        'que será mostrado ao usuário caso haja um erro de integridade referencial
                                        'com a tabela (duplicidade, etc...)
                                        
                                        'Campo na tabela do tipo texto, identificado por "str"
                                        If UCase(Mid(.Controls(i).Name, 4, 3)) = "STR" Then
                                            vetCampos(1, j) = .Controls(i).Text
                                        Else
                                            vetCampos(1, j) = .Controls(i).BoundText & "_" & .Controls(i).Text
                                        End If
                                    Else
                                        'Campo na tabela do tipo texto, identificado por "str"
                                        If UCase(Mid(.Controls(i).Name, 4, 3)) = "STR" Then
                                            'Guarda o valor preenchido (Text) do DataCombo
                                            vetCampos(1, j) = Trim(.Controls(i).Text)
                                        'Campo na tabela do tipo numérico, identificado por "int"
                                        ElseIf UCase(Mid(.Controls(i).Name, 4, 3)) = "INT" Then
                                            'Atribui-se um caracter junto a descrição do campo para que possa, no COM/DCOM,
                                            'padronizar mensagens de erro de conversão de tipos de dados.
                                            '(O usuário poderá digitar um número, aí teríamos outro tipo de erro como
                                            'Constraint, integridade de relacionamentos, duplicidade, etc...)
                                            'SELECT PKId, strDescricao FROM tblBairro ORDER BY strDescricao;strDescricao
                                            If .Controls(i).Tag <> "" Then
                                                strSql = Left(.Controls(i).Tag, InStr(.Controls(i).Tag, ",") - 1)
                                                
                                                If Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), " ")) = "" Then
                                                    strSql = strSql & " FROM " & Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), Len(.Controls(i).Tag))
                                                Else
                                                    strFinalString = Right(Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), " ") + 5), 6)
                                                    Select Case Trim(UCase(strFinalString))
                                                        Case "ORDER", "GROUP", "WHERE"
                                                            strSql = strSql & " FROM " & Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), " "))
                                                        Case Else
                                                            If Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), ",")) <> "" Then
                                                                strSql = strSql & " FROM " & Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl")), ","))
                                                            ElseIf Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), ";")) <> "" Then
                                                                strSql = strSql & " FROM " & Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), InStr(Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStr(.Controls(i).Tag, "tbl") + 1), ";"))
                                                            Else
                                                                strSql = strSql & " FROM " & Mid(.Controls(i).Tag, InStr(.Controls(i).Tag, "tbl"), Len(.Controls(i).Tag))
                                                            End If
                                                    End Select
                                                End If
                                                strSql = strSql & " WHERE " & Right(.Controls(i).Tag, Len(.Controls(i).Tag) - InStrRev(.Controls(i).Tag, ";"))
                                                strSql = strSql & " ='" & .Controls(i).Text & "'"
                        
                                                'On Error Resume Next
                                                strLabel = .Controls(Replace(.Controls(i).Name, "dbc", "lbl")).Caption & " "
                                                'On Error GoTo 0
                                                'EXIT FUNCTION
                                                
                                            Else
                                                'On Error Resume Next
                                                strLabel = .Controls(Replace(.Controls(i).Name, "dbc", "lbl")).Caption & " "
                                                'On Error GoTo 0
                                                            
                                                If .Controls(i).Enabled Then
                                                    .Controls(i).SetFocus
                                                End If
                                                SalvarOperacaoUsuario = False
                                                'EXIT FUNCTION
                                            'vetCampos(1, j) = "[" & Trim(.Controls(i).Text) & "]"
                                            End If
                                        Else
                                            'EXIT FUNCTION
                                        End If
                                    End If
                                ElseIf TypeOf .Controls(i) Is ComboBox Then
                                    If .Controls(i).ListIndex >= 0 Then
                                        'Campo na tabela do tipo texto, identificado por "str"
                                        If UCase(Mid(.Controls(i).Name, 4, 3)) = "STR" Then
                                            vetCampos(1, j) = .Controls(i).Text
                                        Else
                                            vetCampos(1, j) = .Controls(i).ItemData(.Controls(i).ListIndex) & "_" & .Controls(i).Text
                                        End If
                                    Else
                                        'Campo na tabela do tipo texto, identificado por "str"
                                        If UCase(Mid(.Controls(i).Name, 4, 3)) = "STR" Then
                                            'Guarda o valor preenchido (Text) do ComboBox
                                            vetCampos(1, j) = .Controls(i).Text
                                        ElseIf UCase(Mid(.Controls(i).Name, 4, 3)) = "INT" Then
                                            'Atribui-se um caracter junto a descrição do campo para que possa, no COM/DCOM,
                                            'padronizar mensagens de erro de conversão de tipos de dados.
                                            '(O usuário poderá digitar um número, aí teríamos outro tipo de erro como
                                            'Constraint, integridade de relacionamentos, duplicidade, etc...)
                                            vetCampos(1, j) = "[" & Trim(.Controls(i).Text) & "]"
                                        Else
                                            'EXIT FUNCTION
                                        End If
                                    End If
                                Else
                                    If UCase(Left(strAux, 3)) = "DTM" Then
                                        If gblnDataValida(.Controls(i), True) Then
        '                                    vetCampos(1, j) = Format(.Controls(i), "YYYY/MM/DD HH:MM:SS")
                                            If bytDBType = EDatabases.SQLServer Then
                                                vetCampos(1, j) = Format(.Controls(i), "YYYY/MM/DD HH:MM:SS")
                                            ElseIf bytDBType = EDatabases.Oracle Then
                                                vetCampos(1, j) = gstrFormataDataOracle(.Controls(i))
                                            End If
                                        Else
                                            'EXIT FUNCTION
                                        End If
                                    ElseIf InStr(UCase(.Controls(i).Name), "INTCEP") > 0 Then
                                        vetCampos(1, j) = gstrValorSemMascara(.Controls(i))
                                    ElseIf InStr(UCase(.Controls(i).Name), "CNPJCPF") > 0 Or _
                                       InStr(UCase(.Controls(i).Name), "CNPJ") > 0 Then
                                        If Len(Trim(gstrValorSemMascara(.Controls(i)))) = 11 Then
                                            If gblnCPFOk(gstrValorSemMascara(.Controls(i))) Then
                                                vetCampos(1, j) = gstrValorSemMascara(.Controls(i))
                                            Else
                                                If .Controls(i).Enabled Then
                                                    .Controls(i).SetFocus
                                                End If
                                                'EXIT FUNCTION
                                            End If
                                        ElseIf Len(Trim(gstrValorSemMascara(.Controls(i)))) = 14 Then
                                            If gblnCGCOk(gstrValorSemMascara(.Controls(i))) Then
                                                vetCampos(1, j) = gstrValorSemMascara(.Controls(i))
                                            Else
                                                If .Controls(i).Enabled Then
                                                    .Controls(i).SetFocus
                                                End If
                                                'EXIT FUNCTION
                                            End If
                                        Else
                                            If .Controls(i).Enabled Then
                                                .Controls(i).SetFocus
                                            End If
                                            'EXIT FUNCTION
                                        End If
                                    ElseIf Val(.Controls(i).Tag) = 1 Then
                                    'Tag de objetos com nº 1 => TextBox que utilizam máscara
                                    'específica e guardam na tabela
                                    'o valor Sem a máscara
                                    '17/05/2001
                                        vetCampos(1, j) = gvntConvFormatoEspecificoParaSQL(.Controls(i))
                                    ElseIf IsNumeric(.Controls(i)) And _
                                           UCase(.Controls(i).Name) <> "TXTSTRMASCARA" And _
                                           Val(.Controls(i).Tag) <> 2 Then
                                        vetCampos(1, j) = gstrConvVrParaSql(.Controls(i))
                                    Else
                                        vetCampos(1, j) = .Controls(i)
                                    End If
                                End If
                            End If
                            
                            'Título do campo
                            '(Elimina o PKID - autonumeração - que será utilizado apenas nas cláusulas WHERE
                            'das query's de ação (não há título para este campo))
                            If UCase(.Controls(i).Name) <> "TXTPKID" Then
                                If TypeOf .Controls(i) Is OptionButton Then
                                    strAux = "opt" & strAux
                                    vetCampos(2, j) = .Controls(i).Caption
                                ElseIf TypeOf .Controls(i) Is CheckBox Then
                                    strAux = "chk" & strAux
                                    vetCampos(2, j) = .Controls(i).Caption
                                Else
                                    strAux = "lbl" & strAux
                                    vetCampos(2, j) = .Controls(strAux).Caption
                                    If blnErro Then
                                        strAux = Replace(strAux, "lbl", "fra_")
                                        vetCampos(2, j) = .Controls(strAux).Caption
                                    End If
                                End If
                            End If
                            'Prefixo do tipo do objeto
                            vetCampos(3, j) = Left(.Controls(i).Name, 3)
                            
                            'Incrementa o índice das linhas do vetor
                            j = j + 1
                        End If
                    End If
                    
                Next i
            End If
        End With
    'gcncADOMain.BeginTrans
    
    strAux = MontaRegistroAlteracaoUsuario(vetCampos, strTabela, strModoOperacao)
    GravaHistoricoOperacaoUsuario strTabela, strModoOperacao, strAux
    
    'Diferente de "" a manutenção no banco de dados não foi bem sucedida
    If strAux <> "" Then
        SalvarOperacaoUsuario = False
        'gcncADOMain.RollbackTrans
    Else
        SalvarOperacaoUsuario = True
    '    gcncADOMain.CommitTrans
    End If

'Erro imprevisto (DEFINIR MELHOR MENSAGEM)
End Function

Private Function MontaRegistroAlteracaoUsuario(vetCampos, strTabela As String, strModoOperacao As String) As String
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
        
        If gobjBanco.CriaADO(strSql, 10, adoRec) Then
            For Each objCampo In adoRec.Fields
                If UCase(objCampo.Name) <> "DTMDTATUALIZACAO" And UCase(objCampo.Name) <> "LNGCODUSR" And _
                   UCase(objCampo.Name) <> "PKID" Then
                    For i = 0 To UBound(vetCampos, 2)
                        If UCase(objCampo.Name) = UCase(vetCampos(0, i)) Then
                            If Not IsNull(objCampo.Value) Then
                                strCampo = objCampo.Value
                            Else
                                strCampo = ""
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
                strAux = "ID = lngPKId"
            End If
        End If
    Case "I" 'Inclusão
        strAux = ""
        For i = 1 To UBound(vetCampos, 2) - 1
            strAux = strAux & vetCampos(2, i) & " = " & Replace(Trim(vetCampos(1, i)), "'", Chr(207)) & Chr(13)
        Next i
End Select

MontaRegistroAlteracaoUsuario = strAux
End Function


Public Sub GravaHistoricoOperacaoUsuario(strTabela As String, strModoOperacao As String, strAux As String)

'******************************************************************************************
' Data: 06/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data          : 31/03/2006
' Alteração     : Troca da declaração da função de Private para Public, para possibilitar a
'               : utilização em rotinas externas ao módulo
' Responsável   : Fernando Peixoto
' Pendência     : Mat0376 (11603)
'******************************************************************************************

Dim strSql As String
Dim adoRec As ADODB.Recordset
Dim lngPKIdTabela As Long

strSql = ""
strSql = strSql & " SELECT PKID FROM " & gstrCatalogoTabela & " WHERE strTabela = '" & strTabela & "'"

Set gobjBanco = New clsBanco
If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    If Not (adoRec.BOF And adoRec.EOF) Then
        lngPKIdTabela = adoRec!Pkid
    Else
        Exit Sub
    End If
End If

strSql = ""
strSql = strSql & " INSERT INTO " & gstrHistoricoOperacao
strSql = strSql & " (intUsuario, intCatalogoTabela, bytModulo, strOperacao, dtmData, strValor) VALUES ("
strSql = strSql & glngCodUsr
strSql = strSql & ", " & lngPKIdTabela
strSql = strSql & ", " & bytRetornaCodigoModulo(App.ProductName)
strSql = strSql & ", '" & strModoOperacao & "'"
'strSql = strSql & ", GETDATE()"
strSql = strSql & ", " & strGETDATE
strSql = strSql & ", '" & strAux & "'"
strSql = strSql & ")"

Set gobjBanco = New clsBanco
gobjBanco.Execute strSql
End Sub

Public Function gblnGerouReduzidoReceita(intExercicio As Integer) As Boolean
    Dim intCodigo       As Integer
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    On Error GoTo ErroReduzidoReceita
    strSql = ""
    strSql = strSql & "SELECT PR.PKId, CO.strCodigoOrcamentario FROM "
    strSql = strSql & gstrPrevisaoDaReceita & " PR, "
    strSql = strSql & gstrCodigoOrcamentario & " CO "
    strSql = strSql & "WHERE PR.intCodigoOrcamentario = CO.PKId "
    strSql = strSql & "AND PR.intExercicio = " & intExercicio & " "
    strSql = strSql & "ORDER BY CO.strCodigoOrcamentario"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                intCodigo = intCodigo + 1
                strSql = ""
                strSql = strSql & "UPDATE " & gstrPrevisaoDaReceita & " "
                strSql = strSql & "SET intCodigoReduzido = " & intCodigo & " "
                strSql = strSql & "WHERE PKId = " & !Pkid
                Set gobjBanco = New clsBanco
                If gobjBanco.Execute(strSql) = False Then
                    Exit Function
                End If
                .MoveNext
            Loop
        End With
    End If
    gblnGerouReduzidoReceita = True
    
ErroReduzidoReceita:
    Resume FimReduzidoReceita
    
FimReduzidoReceita:
End Function

Public Sub LeCodigoOrcamentarioGeral(cboCodigo As ComboBox, _
                                     cboDescricao As ComboBox, _
                            Optional strQuery As String)
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    cboCodigo.Clear
    cboDescricao.Clear
    If Trim(strQuery) = "" Then
        strSql = ""
        strSql = strSql & "SELECT CO.PKId, CO.strCodigoOrcamentario, CO.strDescricao "
        strSql = strSql & "FROM "
        strSql = strSql & gstrCodigoOrcamentario & " CO "
    Else
        strSql = strQuery
    End If
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                cboDescricao.AddItem !strDescricao
                cboDescricao.ItemData(cboDescricao.NewIndex) = !Pkid
                cboCodigo.AddItem gvntFormatacaoEspecifica(!strCodigoOrcamentario)
                cboCodigo.ItemData(cboCodigo.NewIndex) = !Pkid
                .MoveNext
            Loop
        End With
    End If
End Sub

Public Function CarregaInstrucoesParcelas(blnFebraban As Boolean, strComposicaoDaReceita As String, intExercicio As Integer, intParcela As Integer, blnParcelaValida As Boolean, lngLancamentoAlfa As Long)
Dim adoResultado    As New ADODB.Recordset
Dim strSql          As String
Dim strInstrucoes   As String
    
    'Estas informacoes sao somente para ficha compensacao
    If Not blnFebraban Then
        strInstrucoes = "Finalidade: " & strComposicaoDaReceita & _
                        " Exercício: " & intExercicio & _
                        " - Parcela: " & intParcela
    End If
    
    strSql = ""
    strSql = strSql & "SELECT Max(LV.intParcela) intParcela, PA.strParcela, PA.strParcelaOpcional "
    strSql = strSql & "FROM "
    strSql = strSql & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA, " & gstrParametroAtualizacao & " PA "
    strSql = strSql & "WHERE LV.intLancamentoAlfa = " & lngLancamentoAlfa & " And LA.Pkid = LV.intLancamentoAlfa AND PA.intComposicaoReceita = LA.intComposicaoDaReceita AND PA.intExercicio = LA.intExercicio AND bitParcelaValida = 1 "
    strSql = strSql & "GROUP BY PA.strParcela, PA.strParcelaOpcional"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 25, adoResultado) Then
        With adoResultado
            If .EOF = False Then
            
                'Vamos concatenar as instrucoes de acordo com o tipo da parcela
                If blnParcelaValida Then
                    If Not blnFebraban Then
                        strInstrucoes = strInstrucoes & "/" & adoResultado("intParcela").Value
                    End If
                    strInstrucoes = strInstrucoes & Chr(13) & Chr(10) & adoResultado("strParcela").Value
                Else
                    strInstrucoes = strInstrucoes & Chr(13) & Chr(10) & adoResultado("strParcelaOpcional").Value
                End If
                
            End If
        End With
        adoResultado.Close: Set adoResultado = Nothing
    End If
    
    CarregaInstrucoesParcelas = strInstrucoes
    
End Function

Public Function gstrLimitaCampoValor(objControle As Object, KeyAscii As Integer, QtdInteiro As Integer, QtdDecimal As Integer) As String

    '--------------------------------------------------------------------------------
    ' FUNÇÃO USADA PARA LIMITAR A QUANTIDADE DE NUMEROS INTEIROS E CASAS DECIMAIS
    ' DURANTE A DIGITAÇÃO
    '--------------------------------------------------------------------------------
    ' objControle   - Controle de onde está sendo chamada a função
    '                 (ps: O controle deve possuir as propriedades .text e .selstart)
    ' KeyAscii      - Valor ASCII do caracter digitado
    ' QtdInteiro    - Quantidade números que poderão ser digitados antes da vírgula
    ' QtdDecimal    - Quantidade números que poderão ser digitados depois da vírgula
    '--------------------------------------------------------------------------------
    ' Por se tratar de números com casas decimais, utiliza-se a função CaracterValido
    ' com seu parâmetro strtipo = "V"
    '--------------------------------------------------------------------------------
    
    Dim qtdPonto    As Integer
    Dim i           As Integer
    
    qtdPonto = 0

    CaracterValido KeyAscii, "V", objControle
           
    If KeyAscii = 44 Then
    If QtdDecimal = 0 Then KeyAscii = 0
    Exit Function 'Vírgula
    End If
    With objControle
        'Limpeza do controle caso todo seu conteúdo tenha sido selecionado
        If .SelLength = Len(objControle) Then
            .Text = ""
        End If
        
        If InStr(1, .Text, ",") > 0 Then
            '.text com Vírgula Digitada
            If InStr(1, .Text, ",") - 1 < .SelStart Then
                'Número digitado depois da Vírgula
                If Len(Mid(.Text & Chr(KeyAscii), InStr(1, .Text, ",") + 1, Len(.Text))) > QtdDecimal Then
                    If KeyAscii <> 8 Then KeyAscii = 0
                End If
            Else
            'Número digitado antes da Vírgula
                'Quantidade de pontos no número já formatado
                For i = 1 To QtdInteiro
                    If InStr(i, Mid(.Text, 1, Len(.Text)), ".") > 0 Then
                        qtdPonto = qtdPonto + 1
                        i = InStr(i, Mid(.Text, 1, Len(.Text)), ".")
                    End If
                Next
                If InStr(1, .Text, ",") - 1 >= QtdInteiro + qtdPonto Then
                    If KeyAscii <> 8 Then KeyAscii = 0
                End If
            End If
        Else
            '.text sem Vírgula Digitada
            If Len(.Text & Chr(KeyAscii)) > QtdInteiro Then
                If KeyAscii <> 8 Then KeyAscii = 0
            End If
        End If
    End With
    
End Function

Public Function gblnNaoImprimeDataHora() As Boolean
    Dim adoTemp As New ADODB.Recordset
    
    gblnNaoImprimeDataHora = False
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO("SELECT bytNaoImprimeDataHora FROM " & gstrConfiguracaoGeral, 10, adoTemp) Then
        If Not adoTemp.EOF Then
            gblnNaoImprimeDataHora = adoTemp!bytNaoImprimeDataHora
        End If
    End If
    Set gobjBanco = Nothing
    
End Function

Public Function gstrStringCripitografadaSimplificada(ByVal vntTexto As Variant, _
                                      Optional blnCripitogravar As Boolean) As String
    '---------------------------------------------------------------------------------
    ' FUNÇÃO USADA PARA CRIPTOGRAFAR E/OU DESCRIPTOGRAFAR TEXTO.
    ' Criada por Eduardo Carmo em 08/2006 para gerar uma criptografia com menos caracteres
    ' para que o usuário possa informar senhas para bloqueio/desbloqueio do sistema
    '---------------------------------------------------------------------------------
    ' PARÂMETRO:
    ' 1 - vntTexto - Texto a ser cripitografado/descripitografado
    ' 2 - blnCripitogravar - Flag indicando se vai cripitogravar ou
    '                        descripitogravar o texto
    '---------------------------------------------------------------------------------
    Dim strTexto                As String
    Dim intInd                  As Integer
    Dim bytAux                  As Byte
    On Error Resume Next
    
    strTexto = ""
    If blnCripitogravar Then
        For intInd = 1 To Len(Trim(vntTexto))
            bytAux = Asc(Mid$(vntTexto, intInd, 1))
            strTexto = strTexto + Chr(Format(((bytAux + 1) * 2), "000"))
        Next
    Else
        For intInd = 1 To Len(Trim(vntTexto))
            bytAux = Asc(Mid$(vntTexto, intInd, 1))
            strTexto = strTexto + Chr((bytAux / 2) - 1)
        Next
    End If
    
    gstrStringCripitografadaSimplificada = strTexto
End Function

Public Function SaldoContaContabilAtual(intContaContabil As Double, intMes As Integer, intExercicio As Integer, Optional dblValor As Double, Optional strVarDia As String) As Double
    
    
    Dim strSql                     As String
    Dim intCont                    As Integer
    Dim dblSaldoContaContabil      As Double
    Dim dblSaldoContaContabilDia   As Double
    Dim dblSaldoContaContabilIni   As Double
    Dim dblSaldo(1 To 12, 0 To 2)  As Double
    Dim dblSaldoIni                As Double
    Dim dblDebito                  As Double
    Dim dblCredito                 As Double
    Dim adoResultado               As New ADODB.Recordset
    
    
    'Vamos buscar o valor do Saldo Inicial
    
    strSql = "SELECT dblValor dblSaldoDaConta, blnNaturezaDaConta FROM " & gstrPlanoContaSaldo
    strSql = strSql & " WHERE intPlanoConta = " & intContaContabil
    strSql = strSql & " AND intExercicio = " & intExercicio
    
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            dblSaldoIni = IIf(IsNull(adoResultado.Fields("dblsaldodaconta").Value), gstrConvVrDoSql("0", 2, 5), adoResultado.Fields("dblsaldodaconta").Value)
            If Val(adoResultado.Fields("blnNaturezaDaConta")) = 0 Then
                dblSaldoIni = dblSaldoIni * (-1)
            End If
        End If
        adoResultado.Close
    End If
    
    
    SaldoContaContabilAtual = dblSaldoIni
    dblSaldoContaContabilIni = dblSaldoIni
    dblSaldoContaContabil = dblSaldoIni
    
    
    
    strSql = "SELECT " & gstrISNULL("SUM(LC.dblValor)", "0") & " AS dblValor FROM " & gstrLancamentoContabil & " LC, "
    strSql = strSql & gstrProcessoPagamento & " PP "
    strSql = strSql & "WHERE LC.intConta = " & intContaContabil & " AND "
    strSql = strSql & "LC.intProcesso = PP.PKID AND "
    strSql = strSql & "NOT PP.Intlancamentocontabil IS NUll AND "
    strSql = strSql & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio & " AND "
    strSql = strSql & "dtmData <= " & gstrConvDtParaSql(strVarDia) & " AND "
    strSql = strSql & "LC.bytNatureza = 1"
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            dblCredito = adoResultado.Fields("dblValor").Value
        End If
        adoResultado.Close
    End If
    
    
    
    strSql = "SELECT " & gstrISNULL("SUM(LC.dblValor)", "0") & " AS dblValor FROM " & gstrLancamentoContabil & " LC, "
    strSql = strSql & gstrProcessoPagamento & " PP "
    strSql = strSql & "WHERE LC.intConta = " & intContaContabil & " AND "
    strSql = strSql & "NOT PP.Intlancamentocontabil IS NUll AND "
    strSql = strSql & "LC.intProcesso = PP.PKID AND "
    strSql = strSql & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio & " AND "
    strSql = strSql & "dtmData <= " & gstrConvDtParaSql(strVarDia) & " AND "
    strSql = strSql & "LC.bytNatureza = 0"
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            dblDebito = adoResultado.Fields("dblValor").Value
        End If
    End If
    
    dblSaldoContaContabilDia = dblSaldoContaContabilIni + (dblCredito - dblDebito)
    
    
    If Len(strVarDia) <> 0 And Val(gstrConvVrParaSql(dblValor)) = 0 Then
        SaldoContaContabilAtual = dblSaldoContaContabilDia
        Exit Function
    End If
    
    
    
    
    For intCont = 1 To 12
        
        strSql = "SELECT " & gstrISNULL("SUM(LC.dblValor)", "0") & " AS dblValor FROM " & gstrLancamentoContabil & " LC, "
        strSql = strSql & gstrProcessoPagamento & " PP "
        strSql = strSql & "WHERE LC.intConta = " & intContaContabil & " AND "
        strSql = strSql & "LC.intProcesso = PP.PKID AND "
        strSql = strSql & "NOT PP.Intlancamentocontabil IS NUll AND "
        strSql = strSql & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio & " AND "
        strSql = strSql & gstrDATEPART(strMONTH, "dtmData") & " = " & intCont & " AND "
        strSql = strSql & "LC.bytNatureza = 1"
        
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                dblCredito = adoResultado.Fields("dblValor").Value
            End If
            adoResultado.Close
        End If
        
        
        
        strSql = "SELECT " & gstrISNULL("SUM(LC.dblValor)", "0") & " AS dblValor FROM " & gstrLancamentoContabil & " LC, "
        strSql = strSql & gstrProcessoPagamento & " PP "
        strSql = strSql & "WHERE LC.intConta = " & intContaContabil & " AND "
        strSql = strSql & "NOT PP.Intlancamentocontabil IS NUll AND "
        strSql = strSql & "LC.intProcesso = PP.PKID AND "
        strSql = strSql & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio & " AND "
        strSql = strSql & gstrDATEPART(strMONTH, "dtmData") & " = " & intCont & " AND "
        strSql = strSql & "LC.bytNatureza = 0"
        
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                dblDebito = adoResultado.Fields("dblValor").Value
            End If
            adoResultado.Close
        End If
        
        
        
        dblSaldo(intCont, 0) = dblSaldoIni
        dblSaldo(intCont, 1) = dblCredito
        dblSaldo(intCont, 2) = dblDebito
        
        
    Next
    
    For intCont = 1 To intMes
        dblSaldoContaContabil = dblSaldoContaContabil + ((dblSaldo(intCont, 1) - dblSaldo(intCont, 2)))
    Next
    
    SaldoContaContabilAtual = dblSaldoContaContabil
    
    If Val(gstrConvVrParaSql(dblValor)) > 0 Then
        If Val(gstrConvVrParaSql(dblValor)) > Val(gstrConvVrParaSql(dblSaldoContaContabilDia)) Then
            If MsgBox("O saldo disponível para esta movimentação é insuficiente." & vbNewLine & "Você Deseja realizar a movimentação mesmo assim?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                SaldoContaContabilAtual = dblSaldoContaContabil
            Else
                'ExibeMensagem "Saldo insulficiente para este movimento."
                SaldoContaContabilAtual = Empty
            End If
            Exit Function
        Else
            For intMes = intMes + 1 To 12
                
                SaldoContaContabilAtual = SaldoContaContabilAtual + ((dblSaldo(intMes, 2)) - dblSaldo(intMes, 1))
                
                If Val(gstrConvVrParaSql(dblValor)) > Val(gstrConvVrParaSql(SaldoContaContabilAtual)) Then
                    
                    If MsgBox("Saldo bancário é insulficiente nos meses posteriores." & vbNewLine & "Você Deseja realizar a movimentação mesmo assim?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                        SaldoContaContabilAtual = dblSaldoContaContabil
                    Else
                        'ExibeMensagem "Saldo bancário insulficiente nos meses posteriores."
                        SaldoContaContabilAtual = Empty
                    End If
                    Exit Function
                End If
                
            Next
        End If
    End If
    
    SaldoContaContabilAtual = dblSaldoContaContabil
    
End Function

Public Sub preencheDotacaoByCodigo(objCboCodigoReduzido As ComboBox, objCbointProgramaDeTrabalho As ComboBox, Optional strQuery As String)
    
    'Funcão criada por Ulisses em 06/11/03
    'Usada no evento KeyPress e LostFocus de combos que possuam
    'o codigo reduzido de dotação
    'Objetivo: Preencher os dados de dotação buscando pelo codigo reduzido
    
    
    Dim adoTemp As ADODB.Recordset
    Dim strSql As String
    
    If objCboCodigoReduzido.Enabled = False Then Exit Sub
    
    If Trim(objCboCodigoReduzido.Text) = "" Then Exit Sub
    
    If Not IsMissing(strQuery) Then
        strSql = ""
        strSql = strSql & "SELECT PKId, intCodigoReduzido, strCodigo "
        strSql = strSql & "FROM " & gstrProgramaDeTrabalho & " "
        
        strSql = strSql & "WHERE intExercicio=" & CStr(gintExercicio) & " "
    Else
        strSql = strQuery
    End If
    strSql = strSql & "AND intCodigoReduzido = " & Trim(objCboCodigoReduzido.Text)
    
    Set gobjBanco = New clsBanco
    gobjBanco.CriaADO strSql, 5, adoTemp
    
    If Not adoTemp.EOF Then
        If objCboCodigoReduzido.ListCount = 0 Then
            LeProgramaTrabalhoComReduzido objCboCodigoReduzido, objCbointProgramaDeTrabalho, gintExercicio
        End If
        objCboCodigoReduzido.ListIndex = ModGeral.gintIndiceCBO(objCboCodigoReduzido, adoTemp!Pkid)
    Else
        objCbointProgramaDeTrabalho.ListIndex = -1
    End If
    
End Sub


Public Sub LeProgramaTrabalhoComReduzido(cboCodigoReduzido As ComboBox, _
    cboProgramaTrabalho As ComboBox, _
    Optional intExercicio As Integer, _
    Optional strQuery As String)
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    cboProgramaTrabalho.Clear
    cboCodigoReduzido.Clear
    
    If strQuery = "" Then
        strSql = ""
        strSql = strSql & "SELECT PKId, intCodigoReduzido, strCodigo "
        strSql = strSql & "FROM " & gstrProgramaDeTrabalho & " "
        
        If Not intExercicio = 0 Then strSql = strSql & "WHERE intExercicio=" & intExercicio & " "
        
        strSql = strSql & "ORDER BY strCodigo"
    Else
        strSql = strQuery
    End If
    
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                If Not IsNull(!intCodigoReduzido) Then
                    cboProgramaTrabalho.AddItem !strCodigo
                    cboProgramaTrabalho.ItemData(cboProgramaTrabalho.NewIndex) = !Pkid
                End If
                .MoveNext
            Loop
        End With
    End If
    
    If strQuery = "" Then
        strSql = Mid(strSql, 1, Len(strSql) - 18) + "ORDER BY intCodigoReduzido"
    End If
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                If Not IsNull(!intCodigoReduzido) Then
                    cboCodigoReduzido.AddItem gstrENulo(!intCodigoReduzido)
                    cboCodigoReduzido.ItemData(cboCodigoReduzido.NewIndex) = !Pkid
                End If
                .MoveNext
            Loop
        End With
    End If
End Sub

Public Function GerarEmpenhoDeEstorno() As Boolean
    Dim strSql       As String
    Dim adoResultado As New ADODB.Recordset
    
    strSql = "SELECT bytGerarEmpenhoDeEstorno FROM " & gstrConfiguracaoGeral
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                If !bytGerarEmpenhoDeEstorno = 1 Then
                    GerarEmpenhoDeEstorno = True
                End If
            End If
        End With
    End If
End Function

Function CancelarReservaDotacao(lngPkidReserva As Long, strDataCancelamento As String, dblValorCancelamento As Double, strHistoricoCancelamento As String, Optional blnPergunta As Boolean = True) As Boolean
    Dim adoResultado     As ADODB.Recordset
    Dim strSql           As String
    
    'ROTINA UTILIZADA PARA CANCELAR RESERVAS DE DOTAÇÃO TOTAL OU PARCIALMENTE
    If blnPergunta = True Then
        If gblnExclusaoGravacaoOk("I", "Deseja que o Saldo da Reserva volte para? " & vbCrLf & "[SIM ] - Dotação" & vbCrLf & "[NÃO] - Reserva", True) = False Then
            CancelarReservaDotacao = False
            Exit Function
        End If
    End If
    strSql = ""
    strSql = strSql & "INSERT INTO " & gstrReservaDotacaoLiberada & " ("
    strSql = strSql & "intReservaDotacao, intNumero, dtmData, dblValor, "
    strSql = strSql & "strHistorico, dtmDtAtualizacao, lngCodUsr, intFlag"
    strSql = strSql & ") (SELECT "
    strSql = strSql & Val(lngPkidReserva) & ", " & gstrISNULL("MAX(intNumero)", "0") & " + 1, "
    strSql = strSql & gstrConvDtParaSql(strDataCancelamento) & ", "
    strSql = strSql & gstrConvVrParaSql(dblValorCancelamento) & ", "
    strSql = strSql & "'" & strHistoricoCancelamento & "', "
    strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
    strSql = strSql & glngCodUsr & ", 0 "
    strSql = strSql & "FROM " & gstrReservaDotacaoLiberada & " "
    strSql = strSql & "WHERE intReservaDotacao = " & Val(lngPkidReserva) & ")"
    Set gobjBanco = New clsBanco
    If gobjBanco.Execute(strSql) Then
        strSql = "SELECT strSolicitacao,intExercicio FROM tblReservaDotacao WHERE Pkid = " & lngPkidReserva
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            
            If adoResultado!strSolicitacao <> "" And adoResultado!intExercicio <> "" Then
                
                strSql = "UPDATE " & gstrRequisicaoCompras & " SET "
                strSql = strSql & "intRequisicaoComprasSituacoes = (SELECT PKId FROM " & gstrRequisicaoComprasSituacoes & " WHERE bitTipo = 1), "
                strSql = strSql & "intReserva = NULL "
                strSql = strSql & "WHERE intCodigo = " & adoResultado!strSolicitacao & " AND intExercicio = " & adoResultado!intExercicio
                
                If gobjBanco.Execute(strSql) Then
                    CancelarReservaDotacao = True
                End If
            End If
            
        End If
        CancelarReservaDotacao = True
    End If
    
End Function

Public Sub TrocaCorDeFundoObjeto(blnAlterando As Boolean, _
    ParamArray vntObjeto())
    Dim objControle As Variant
    For Each objControle In vntObjeto
        If blnAlterando Then
            If UCase(Mid(objControle.Name, 1, 4)) <> "CMD_" Then
                objControle.BackColor = Val(gvntFundoObjInacessivel)
            End If
            objControle.Enabled = False
        ElseIf objControle.OLEDropMode = 0 Then
            If UCase(Mid(objControle.Name, 1, 4)) <> "CMD_" Then
                objControle.BackColor = vbWindowBackground
            End If
            objControle.Enabled = True
        Else
            objControle.BackColor = Val(gvntFundoObjInacessivel)
        End If
    Next
End Sub


Public Sub LeProgramaTrabalho(cboPrograma As Object, cboReduzido As Object, _
    objOrgao As Object, objSubunidade As Object, _
    objFuncao As Object, objPrograma As Object, _
    objProjetoAtividade As Object, objUnidade As Object, _
    objTipoCredito As Object, objSubfuncao As Object, _
    objSubPrograma As Object, objElemento As Object, _
    objSaldo As Object, objTotalDotado As Object, _
    objValor As Object, _
    Optional objBloqueado As Object, _
    Optional objSuplementado As Object, _
    Optional objValorReduzido As Object, _
    Optional objReservado As Object, _
    Optional objCodigoReduzido As Object, _
    Optional objCodigo As Object, _
    Optional objVinculo As Object, _
    Optional objFonte As Object, _
    Optional objGrupo As Object, _
    Optional strData As String)
    
    '******************************************************************************************
    ' Data: 09/06/2003
    ' Alteração: - Substituição da chamada direta à stored procedure pela função
    '            gstrStoredProcedure.
    ' Responsável: Everton Bianchini
    '******************************************************************************************
    
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    Screen.MousePointer = vbHourglass
    strSql = ""
    '    strSql = strSql & "sp_ProgTrabalhoParaEmpnho "
    If TypeOf cboPrograma Is ComboBox Then
        cboReduzido.ListIndex = gintIndiceCBO(cboReduzido, gstrItemData(cboPrograma))
        strSql = strSql & gstrItemData(cboPrograma)
    Else
        strSql = strSql & cboPrograma.Columns(0).Value
    End If
    
    strSql = gstrStoredProcedure("sp_ProgTrabalhoParaEmpnho", strSql, True)
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                objOrgao = gstrENulo(!STRORGAO)
                objSubunidade = gstrENulo(!strSubUnidade)
                objFuncao = gstrENulo(!strFuncao)
                objPrograma = gstrENulo(!strPrograma)
                objProjetoAtividade = gstrENulo(!STRPROJETO)
                objUnidade = gstrENulo(!strUnidadeOrcamentaria)
                objTipoCredito = gstrENulo(!strTipoCredito)
                objSubfuncao = gstrENulo(!strSubFuncao)
                objSubPrograma = gstrENulo(!strSubprograma)
                objElemento = gstrENulo(!strElemento)
                'objSaldo = gstrConvVrDoSql(!dblSaldo)
                If objCodigoReduzido Is Nothing = False Then
                    objCodigoReduzido = !intCodigoReduzido
                End If
                If objCodigo Is Nothing = False Then
                    objCodigo = !strCodigo
                End If
                If objVinculo Is Nothing = False Then
                    objVinculo = gstrENulo(!strVinculo)
                End If
                If objFonte Is Nothing = False Then
                    objFonte = gstrENulo(!strFonte)
                End If
                If objGrupo Is Nothing = False Then
                    objGrupo = gstrENulo(!strGrupo)
                End If
            Else
                objOrgao = ""
                objSubunidade = ""
                objFuncao = ""
                objPrograma = ""
                objProjetoAtividade = ""
                objUnidade = ""
                objTipoCredito = ""
                objSubfuncao = ""
                objSubPrograma = ""
                objElemento = ""
                objSaldo = ""
                objTotalDotado = ""
                objValor = ""
                If objBloqueado Is Nothing = False Then
                    objBloqueado = ""
                End If
                If objSuplementado Is Nothing = False Then
                    objSuplementado = ""
                End If
                If objValorReduzido Is Nothing = False Then
                    objValorReduzido = ""
                End If
                If objReservado Is Nothing = False Then
                    objReservado = ""
                End If
                If objCodigoReduzido Is Nothing = False Then
                    objCodigoReduzido = ""
                End If
                If objCodigo Is Nothing = False Then
                    objCodigo = ""
                End If
                If objVinculo Is Nothing = False Then
                    objVinculo = ""
                End If
                If objFonte Is Nothing = False Then
                    objFonte = ""
                End If
                If objGrupo Is Nothing = False Then
                    objGrupo = ""
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
                
            End If
        End With
    End If
    
    
    If Len(Trim(strData)) > 0 Then
        
        strSql = ""
        strSql = strSql & gstrConvDtParaSql(CStr(gintExercicio) + "/01/01") & ", " & gstrConvDtParaSql(strData) & ", " & gstrItemData(cboPrograma)
        strSql = gstrStoredProcedure("sp_ContaValAcumuladoDia", strSql, True)
        
        'strSql = strSql & " SELECT saldoIni, SUM(Empenhado) dblEmpenhado , "
        'strSql = strSql & " SUM(Suplementado) dblSuplementado,"
        'strSql = strSql & " SUM(Anulado) dblAnulado,SUM(Bloqueado) dblBloqueado,"
        'strSql = strSql & " SUM(Reservado) dblReservado"
        'strSql = strSql & " FROM " & gstrContaValoresAcumulados & " "
        'strSql = strSql & " WHERE intProgramadeTrabalho = " & gstrItemData(cboPrograma)
        'strSql = strSql & " AND intExercicio = " & gintExercicio
        'strSql = strSql & " AND Mes <= " & Month(CDate(strData))
        'strSql = strSql & " GROUP BY saldoIni"
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
            With adoResultado
                
                If Not .EOF Then
                    'objSaldo = gstrConvVrDoSql(SaldoDotacaoAtual(gstrItemData(cboPrograma), Val(Month(CDate(strData))), gintExercicio))
                    
                    objSaldo = gstrConvVrDoSql((!dblSaldoIni + !dblSuplementado) - !dblAnulado - !dblEmpenhado - !dblBloqueado - !dblReservado)
                    
                    objTotalDotado = gstrConvVrDoSql(!dblEmpenhado)
                    
                    'objValor = gstrConvVrDoSql(!DBLVALOR)
                    objValor = gstrConvVrDoSql(!dblSaldoIni)
                    
                    If objBloqueado Is Nothing = False Then
                        objBloqueado = gstrConvVrDoSql(!dblBloqueado)
                    End If
                    If objSuplementado Is Nothing = False Then
                        objSuplementado = gstrConvVrDoSql(!dblSuplementado)
                    End If
                    If objValorReduzido Is Nothing = False Then
                        objValorReduzido = gstrConvVrDoSql(!dblAnulado)
                    End If
                    If objReservado Is Nothing = False Then
                        objReservado = gstrConvVrDoSql(!dblReservado)
                    End If
                End If
                
            End With
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Public Sub LePlanoContaGeral1(ParamArray Parametro())
    Dim vntItem             As Variant
    Dim strAux              As String
    Dim strSql              As String
    Dim strCondicao         As String
    Dim blnConcatenaItem    As Boolean
    Dim objLista()          As Object
    Dim bytInd              As Byte
    Dim bytLimiteObj        As Byte
    Dim adoResultado        As ADODB.Recordset
    strAux = ""
    For Each vntItem In Parametro
        If TypeOf vntItem Is ComboBox Then
            bytInd = bytInd + 1
            bytLimiteObj = bytInd
            ReDim Preserve objLista(bytInd)
            Set objLista(bytInd) = vntItem
            objLista(bytInd).Clear
        ElseIf InStr(UCase(vntItem), "ED") <> 0 Then
            strAux = ""
            strAux = strAux & "SELECT PC.PKId, PC.strContaContabil, PC.strDescricao "
            strAux = strAux & "FROM " & gstrEventoContaContabilDebito & " ED,"
            strAux = strAux & gstrPlanoConta & " PC"
            strAux = strAux & " WHERE "
            strAux = strAux & "ED.intContaContabil = PC.PKID AND "
            strAux = strAux & "ED.intEvento = " & Mid(vntItem, 3)
        ElseIf InStr(UCase(vntItem), "EC") <> 0 Then
            strAux = ""
            strAux = strAux & "SELECT PC.PKId, PC.strContaContabil, PC.strDescricao "
            strAux = strAux & "FROM " & gstrEventoContaContabilCredito & " EC,"
            strAux = strAux & gstrPlanoConta & " PC"
            strAux = strAux & " WHERE "
            strAux = strAux & "EC.intContaContabil = PC.PKID AND "
            strAux = strAux & "EC.intEvento = " & Mid(vntItem, 3)
        ElseIf InStr(UCase(vntItem), "PD") <> 0 Then 'Contas Patrimoniais
        strAux = ""
        strAux = strAux & "SELECT PC.PKId, PC.strContaContabil, PC.strDescricao "
        strAux = strAux & "FROM " & gstrEventoContaContabilDebito & " ED,"
        strAux = strAux & gstrPlanoConta & " PC"
        strAux = strAux & " WHERE "
        strAux = strAux & "ED.intContaContabil = PC.PKID AND "
        strAux = strAux & "ED.intEvento = " & Mid(vntItem, 3) & " AND "
        strAux = strAux & "PC.blnPatrimonial = 1 "
    ElseIf InStr(UCase(vntItem), "PC") <> 0 Then 'Contas Patrimoniais
    strAux = ""
    strAux = strAux & "SELECT PC.PKId, PC.strContaContabil, PC.strDescricao "
    strAux = strAux & "FROM " & gstrEventoContaContabilCredito & " EC,"
    strAux = strAux & gstrPlanoConta & " PC"
    strAux = strAux & " WHERE "
    strAux = strAux & "EC.intContaContabil = PC.PKID AND "
    strAux = strAux & "EC.intEvento = " & Mid(vntItem, 3) & " AND "
    strAux = strAux & "PC.blnPatrimonial = 1 "
    
ElseIf InStr(UCase(vntItem), "SELECT") = 0 Then
    If Trim(strAux) = "" Then
        strAux = strAux & " WHERE "
    ElseIf blnConcatenaItem Then
        If Trim(UCase(strCondicao)) = "OU" Then
            strAux = strAux & " OR "
        Else
            strAux = strAux & " AND "
        End If
    End If
    If Trim(vntItem) = "OU" Then
        blnConcatenaItem = False
    Else
        blnConcatenaItem = True
    End If
Else
    Select Case UCase(vntItem)
    Case "FN" 'Financeira
        strAux = strAux & "ABS(PC.blnFinanceira) = 1"
    Case "RT" 'Retenção
        strAux = strAux & "ABS(PC.blnRetencao) = 1"
    Case "EO" 'Extra-orçamentária
        strAux = strAux & "ABS(PC.blnExtraOrcamentaria) = 1"
    Case "IB" 'Integrar balanço
        strAux = strAux & "ABS(PC.blnIntegraBalanco) = 1"
    Case "RF" 'Retificadora
        strAux = strAux & "ABS(PC.blnRetificadora) = 1"
    Case "IS" 'Inversão de saldo
        strAux = strAux & "ABS(PC.blnInversaoDeSaldo) = 1"
    Case "CR" 'Conta de natureza credora
        strAux = strAux & "ABS(PC.blnNaturezaDaConta) = 0"
    Case "DV" 'Conta de natureza devedora
        strAux = strAux & "ABS(PC.blnNaturezaDaConta) = 1"
    Case "PA"
        strAux = strAux & "ABS(PC.blnPatrimonial) = 1"
    Case "OU"
        strCondicao = Trim(UCase(vntItem))
    Case Else
        strAux = strAux & vntItem
    End Select
End If
Next

If Trim(strAux) <> "" Then
    If Mid(strAux, 1, 6) <> "SELECT" Then
        strAux = strAux & " AND blnAnalitica = 1 "
    End If
Else
    strAux = strAux & "WHERE blnAnalitica = 1 "
End If

If InStr(UCase(strAux), "SELECT") <> 0 Then
    strSql = strAux
Else
    strSql = ""
    strSql = strSql & "SELECT PC.PKId, PC.strContaContabil, PC.strDescricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrPlanoConta & " PC "
    strSql = strSql & strAux & " ORDER BY PC.strDescricao"
End If
Set gobjBanco = New clsBanco
If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
    With adoResultado
        Do While .EOF = False
            For bytInd = 1 To bytLimiteObj
                If bytInd Mod 2 = 0 Then
                    objLista(bytInd).AddItem !strDescricao
                    objLista(bytInd).ItemData(objLista(bytInd).NewIndex) = !Pkid
                Else
                    objLista(bytInd).AddItem gvntFormatacaoEspecifica(!strContaContabil)
                    objLista(bytInd).ItemData(objLista(bytInd).NewIndex) = !Pkid
                End If
            Next
            .MoveNext
        Loop
    End With
End If
End Sub

