IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoIPTU' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoIPTU
GO 
-- sp_CalculoFormulaExecutada -24,NULL, '"15,65,67,62,68,97","01010010015001","01010010015001",16,0,3,2001,8,10,45' 

-- SELECT * FROM tblImobiliario 
-- SELECT * FROM tblReceita
CREATE PROCEDURE sp_CalculoIPTU(@strReceitas		AS NVARCHAR(4000),		
				@strInscricaoInicial	AS NVARCHAR(50),
				@strInscricaoFinal	AS NVARCHAR(50),
				@intComposicaoDaReceita AS INT,
			        @intParcelaInicial      AS INT ,
			        @intParcelaFinal        AS INT ,	
				@intExercicio		AS INT,
				@intOcorrencia		AS INT,
				@dblDesconto		AS MONEY,
				@glngCodUsr		AS INT)
AS
	DECLARE @strQueryLoop		 NVARCHAR(1000),--MONTA STRING A FAZER LOOP 
		@PKId 			 INT,     	--PKID DO IMÓVEL ATUAL DO LOOP
		@bytEdificado		 TINYINT, 	--FLAG QUE INDICA SE É EDIFICADO
		@strInscricalAtual	 NVARCHAR(30), 	--INSCRIÇÃO CADASTRAL ATUAL
		@intUtilizacaoOcorrencia INT,	  	--UTILIZAÇÃO DA OCORRÊNCIA
		@intCodigoOcorrencia	 INT,	  	--CÓDIGO DA OCORRÊNCIA
		@intReceita		 INT,		--AUXILIAR PARA MONTAR AS RECEITAS	
		@strParametros		 NVARCHAR(100), --GUARDA PARÂMETROS DE SP'S
		@dblValorVenalTerreno    MONEY,	  	--VALOR VENAL DO TERRENO
		@dblValorVenalEdificacao MONEY,	  	--VALOR VENAL DA EDIFICAÇÃO
		@dblValorVenalImovel     MONEY,	  	--VALOR VENAL DO IMÓVEL
		@intQuantidadeParcela    INT,		--QUANTIDADE DE PARCELAS 
		@dtmDtLancamento	 DATETIME,	--DATA DO LANÇAMENTO CALCULO
		@dtmDtVencimento	 DATETIME,	--DATA DO 1º VENCIMENTO
		@intContribuinte	 INT,		--PKId DO CONTRIBUINTE
		@strContribuinte	 NVARCHAR(400)	--PKId + INSCRICAO 

	--UTILIZADO PARA MONTAR STRING DAS RECEITAS A SEREM CALCULADAS
	CREATE TABLE #t_Receitas
		     (PKId 	 INT,
		      blnFixo	 BIT,
		      blnValido	 BIT)
	EXECUTE('INSERT INTO #t_Receitas
		 SELECT PKid, 0, 0 FROM tblReceita
			 WHERE PKId IN('+@strReceitas+')')
	UPDATE #t_Receitas SET blnFixo = 1, blnValido = 1
	WHERE PKId IN (15,65,68,97)

	--INICIALIZAÇÃO DE VARIÁVEIS------------------------------------------------------
	SET @intUtilizacaoOcorrencia = 6
	SET @intCodigoOcorrencia = 1	
	SET @strQueryLoop = ''
	IF @strInscricaoInicial != '' AND @strInscricaoFinal != ''
		IF @strInscricaoInicial = @strInscricaoFinal
			SET @strQueryLoop =  ' AND IM.strInscricaoAnterior = "' + @strInscricaoInicial  + '"'
		ELSE
			SET @strQueryLoop =  ' AND IM.strInscricaoAnterior BETWEEN "' 
						+ @strInscricaoInicial + '" AND "' + @strInscricaoFinal + '"'
	--FIM DA INICIALIZAÇÃO------------------------------------------------------------

	SET @strQueryLoop =  'DECLARE c_Imovel CURSOR FOR
				SELECT IM.PKId, IM.bytEdificado, IM.strInscricaoAnterior
				  FROM tblImobiliario IM,tblContribuinte CC,tblOcorrencia OC  
			         WHERE CC.PKId = IM.intContribuinte  
		  		   AND OC.PKId = IM.intOcorrrencia  
		  		   AND OC.intUtilizacaoDaOcorrencia = ' + CONVERT(NVARCHAR,@intUtilizacaoOcorrencia) + 
		  		 ' AND OC.intCodigo = ' + CONVERT(NVARCHAR,@intCodigoOcorrencia) +
	       		    ' AND IM.intComposicao = ' + CONVERT(NVARCHAR,@intComposicaoDaReceita) +
				@strQueryLoop +
			    ' ORDER BY CONVERT(NUMERIC,strInscricaoAnterior) '

	EXECUTE (@strQueryLoop)
	OPEN	c_Imovel
	FETCH	c_Imovel INTO
		@PKId, @bytEdificado, @strInscricalAtual
	WHILE @@FETCH_STATUS = 0
	BEGIN
		SET @strParametros = ''+ CONVERT(NVARCHAR(30),@PKId) +', @dblValor OUTPUT'
		--VERIFICA SE O IMÓVEL É EDIFICADO---------------------------------------------
		IF @bytEdificado = 1 
		BEGIN
			UPDATE #t_Receitas SET blnValido = 1
		 	 WHERE PKId = 101 --IPU
			--VALOR VENAL EDIFICACAO
			EXECUTE sp_CalculoFormulaExecutada -2, @dblValorVenalEdificacao OUTPUT, 
							   @strParametros
		END
		ELSE    
			UPDATE #t_Receitas SET blnValido = 1
		 	 WHERE PKId = 67 --ILIMINAÇÃO PÚBLICA

		--USADA PARA MONTAR(CONCATENAR) RECEITAS A SEREM CALCULADAS
		SET @strReceitas = ''
		DECLARE c_MontaReceitas CURSOR FOR
			SELECT PKId FROM #t_Receitas
			 WHERE blnValido = 1
		OPEN	c_MontaReceitas
		FETCH	c_MontaReceitas INTO
			@intReceita
		WHILE @@FETCH_STATUS = 0
		BEGIN
			SET @strReceitas = @strReceitas + CONVERT(NVARCHAR,@intReceita) + ','
			FETCH c_MontaReceitas INTO
				@intReceita
		END
		CLOSE c_MontaReceitas
		DEALLOCATE c_MontaReceitas
		SET @strReceitas = SUBSTRING(@strReceitas, 1,LEN(@strReceitas)-1)
		UPDATE #t_Receitas SET blnValido = 0
	 	 WHERE blnFixo = 0

		/*--------------------------------------------------------------------------------*/
		--ATUALIZA OS DADOS ENCONTRADOS NO CALCULO DO IPTU (SOMENTE SALVAR NA TABELA)
		/*--------------------------------------------------------------------------------*/
		EXECUTE sp_CalculoFormulaExecutada -1, @dblValorVenalTerreno OUTPUT, @strParametros
		EXECUTE sp_CalculoFormulaExecutada -3, @dblValorVenalImovel OUTPUT, @strParametros

		UPDATE tblImobiliario SET dblValorEdificacao = ISNULL(@dblValorVenalEdificacao,0),
					     dblValorTerreno = ISNULL(@dblValorVenalTerreno,0),
					      dblValorImovel = ISNULL(@dblValorVenalImovel,0),
						   lngCodUsr = @glngCodUsr
					          WHERE PKId = @PKId

		/*-----------------------------------------------------------------------*/
			--GRAVA OS DADOS DO LANÇAMENTO DO IPTU
		/*-----------------------------------------------------------------------*/
		
		SET @intContribuinte = (SELECT intContribuinte FROM tblImobiliario
					 WHERE PKId = @PKId)
		SET @strContribuinte = ('SELECT '+ CONVERT(NVARCHAR(30),@intContribuinte) + 
					',"' + @strInscricalAtual+'"') 
		SET @dtmDtLancamento = (GETDATE())
		SET @dtmDtVencimento = (SELECT dtmDataDaParcela AS dtmDatadeVencimento
					       FROM tblVencimentosDasParcelas VP,
						    tblVencimentos VC
					      WHERE VC.PKId = VP.intVencimento
						AND VC.intTributo = 16
						AND VP.intNumero = 1 --VERIFICAR
						AND YEAR(dtmDataDaParcela) = @intExercicio)

		SET @intQuantidadeParcela = (SELECT COUNT(VP.PKId) AS Quantidade 
					       FROM tblVencimentosDasParcelas VP,
						    tblVencimentos VC
					      WHERE VC.PKId = VP.intVencimento
						AND VP.intNumero BETWEEN @intParcelaInicial AND @intParcelaFinal
						AND VC.intTributo = 16
						AND VP.intExercicio = @intExercicio)	

		--EFETUA INSERÇÕES EM LANÇAMENTO CÁLCULO E PARCELAS TAXAS E RECEITAS-----------------
		EXECUTE sp_CalculoLancamentoReceitas 16, -1, -1, @strReceitas,
						     @strContribuinte, @intExercicio,
						     @intComposicaoDaReceita, NULL, 
						     @dtmDtLancamento,@dtmDtVencimento,
						     @intParcelaInicial,@intParcelaFinal,0,1,
						     @intOcorrencia, 1,	@gLngCodUsr, @PKId, 
						     @dblDesconto		
		-------------------------------------------------------------------------------------
		FETCH c_Imovel INTO
			@PKId, @bytEdificado, @strInscricalAtual
	END
	CLOSE c_Imovel
	DEALLOCATE c_Imovel



