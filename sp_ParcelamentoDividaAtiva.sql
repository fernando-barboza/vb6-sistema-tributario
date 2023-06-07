/*sp_ParcelamentoDividaAtiva (@intContribuinte,	@strInscricaoCadastral,	@intOcorrencia, @intParcelas,
			      @dtmDataVencimento, @intIntervalo, @intDesconto, @strExercicios,
			      @intCodUsuario 		AS INT)*/

IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_ParcelamentoDividaAtiva' AND TYPE = 'P')
   DROP PROCEDURE sp_ParcelamentoDividaAtiva
GO

CREATE PROCEDURE sp_ParcelamentoDividaAtiva(@intContribuinte	   AS INT,
					    @strInscricaoCadastral AS NVARCHAR(30),
					    @intOcorrencia	   AS INT,      
					    @intParcelas 	   AS INT,
					    @dtmDataVencimento 	   AS DATETIME,
					    @intIntervalo 	   AS INT,
					    @intDesconto 	   INT = 0,
					    @strExercicios         AS NVARCHAR(100),
					    @intCodUsuario 	   AS INT)
AS
	DECLARE @strSql				AS NVARCHAR(1000),
		@strCondicao			AS NVARCHAR(500),
		@dblValorTotal		 	AS MONEY,
		@intComposicaoDaReceita  	AS INT,
		@dtmLancamento		 	AS DATETIME,
		@intExercicio		 	AS INT,
		@bitUtilizacaoDebito	 	AS TINYINT,
		@bytOrigem		 	AS TINYINT, 
		@dblValorParcelaReceita  	AS MONEY,
       		@intIndiceParcelaReceita 	AS INT,
		@dblValorDesconto	 	AS MONEY,
		@dblValorDescontoParcela	AS MONEY,
       		@intSequencia            	AS INT,
		@strAuxPKIdReceita 		AS NVARCHAR(10),
		@strPKIdReceita 		AS NVARCHAR(100),
		@dblValorResto		 	AS MONEY,
		@dblValorRestoDesconto 		AS MONEY,
		@intContaPagina			AS INT,
    		@intDividaAtiva			AS INT,
    		@intNumeroPagina		AS INT,
    		@intNumeroInscricao		AS INT,
    		@dtmDataInscricao		AS DATETIME,
    		@intNumeroLivroInscricao	AS INT,
    		@bytResultado			AS TINYINT
        
	SET @strCondicao = ' WHERE PAR.bytAtiva = 1 
				   AND PAR.dtmDataVencimento < GETDATE() 
				   AND LAN.intContribuinte = ' + CONVERT(NVARCHAR, @intContribuinte) + 
				 ' AND LAN.strInscricaoCadastral = "' + @strInscricaoCadastral + 
				'" AND LAN.intExercicio IN (' + @strExercicios + ') 
				   AND LAN.PKId = PAR.intLancamentoCalculo '

	SET @strSql = 'DECLARE c_ParcelaReceita CURSOR 
         	    		FOR
				SELECT SUM(PAR.dblValorParcela) AS dblValorTotal, 
				       PAR.intComposicaoDaReceita, LAN.intExercicio, 
				       LAN.dtmLancamento, LAN.bitUtilizacaoDebito, 
				       LAN.bytOrigem 
				  FROM tblParcelaReceita PAR, tblLancamentoCalculo LAN' 
				 + @strCondicao +
			  	 'GROUP BY PAR.intComposicaoDaReceita, LAN.intExercicio, 
					  LAN.dtmLancamento, LAN.bitUtilizacaoDebito, 
					  LAN.bytOrigem'
	EXECUTE(@strSql)

	OPEN    c_ParcelaReceita
	FETCH   c_ParcelaReceita INTO
		@dblValorTotal, @intComposicaoDaReceita, @intExercicio, 
		@dtmLancamento, @bitUtilizacaoDebito, @bytOrigem
	IF @@CURSOR_ROWS > 0
	BEGIN
		-- DELEÇÃO NA TABELA Parcela Taxa
		SET @strSql = 'DELETE FROM tblParcelaTaxa WHERE intLancamentoCalculo IN 
			      (SELECT DISTINCT PAR.intLancamentoCalculo 
				 FROM tblParcelaReceita PAR, tblLancamentoCalculo LAN'
				+ @strCondicao + ') '

		-- DELEÇÃO NA TABELA Parcela Receita
		SET @strSql = @strSql + 'DELETE FROM tblParcelaReceita WHERE PKId IN 
					(SELECT PAR.PKId
					   FROM tblParcelaReceita PAR, tblLancamentoCalculo LAN'
					+ @strCondicao + ') '

		-- DELEÇÃO NA TABELA Detalhe da Dívida Ativa
		SET @strSql = @strSql + 'DELETE FROM tblDetalheDividaAtiva 
					  WHERE strInscricaoCadastral = "' + @strInscricaoCadastral + 
					 '" AND intDividaAtiva IN 
					(SELECT PKId FROM tblDividaAtiva 
					  WHERE intContribuinte = ' + CONVERT(NVARCHAR, @intContribuinte) + ') ' 
		WHILE @@FETCH_STATUS = 0
		BEGIN

			-- Pesquisa a sequência da composição da receita
			SET @intSequencia = (SELECT ISNULL(MAX(strSequencia),0) + 1 AS Maximo 
					       FROM tblLancamentoCalculo
					      WHERE intComposicaoReceita = @intComposicaoDaReceita
						AND intContribuinte = @intContribuinte
						AND intExercicio = YEAR(GETDATE()))

			-- INSERE LANÇAMENTO CALCULO
			INSERT INTO tblLancamentoCalculo
			           (intExercicio, intContribuinte, intComposicaoReceita, 
				    intMensagem, strInscricaoCadastral,	dtmLancamento, 
				    dtmVencimento, intNumeroDeParcelas, 
				    intIntervaloEntreParcelas, bitUtilizacaoDebito, 
				    intOcorrencia, bytOrigem, strSequencia, dtmDtAtualizacao, 
				    lngCodUsr )	VALUES 
				   (YEAR(GETDATE()), @intContribuinte, @intComposicaoDaReceita,
			            NULL, @strInscricaoCadastral, @dtmLancamento, 
				    @dtmDataVencimento, @intParcelas, @intIntervalo, 
				    @bitUtilizacaoDebito, @intOcorrencia, @bytOrigem, 
				    @intSequencia, GETDATE(), @intCodUsuario) 
			
			DECLARE c_BuscaPKIdReceita CURSOR KEYSET FOR
				SELECT A.PKId FROM tblReceita A, tblValorCompoRec B
				 WHERE A.PKId = B.intReceita
				   AND B.intComposicaoDaReceita = @intComposicaoDaReceita
			SET @strAuxPKIdReceita = ''
			SET @strPKIdReceita = ''
			OPEN    c_BuscaPKIdReceita
			FETCH   c_BuscaPKIdReceita INTO @strPKIdReceita
			IF @@CURSOR_ROWS > 0
			BEGIN
				WHILE @@FETCH_STATUS = 0
				BEGIN
					IF @strAuxPKIdReceita <> '' 
						SET @strPKIdReceita = @strPKIdReceita + ', '
					SET @strPKIdReceita = @strPKIdReceita + @strAuxPKIdReceita
					FETCH c_BuscaPKIdReceita INTO 
						@strAuxPKIdReceita
				END
			END
			CLOSE c_BuscaPKIdReceita
			DEALLOCATE c_BuscaPKIdReceita 

			DECLARE c_BuscaCamposDetalheDividaAtiva CURSOR 
				FOR
				SELECT DIV.PKId, DET.intNumeroPaginaInscricao,
				       DET.intNumeroInscricao, DET.dtmInscricao,
				       DET.intNumeroLivroInscricao
				  FROM tblDetalheDividaAtiva DET, tblDividaAtiva DIV
				 WHERE DIV.intContribuinte = @intContribuinte
				   AND DET.strInscricaoCadastral = @strInscricaoCadastral
				   AND DET.intComposicaoReceita = @intComposicaoDaReceita
				   AND DET.intExercicio = @intExercicio
				   AND DET.intDividaAtiva = DIV.PKId
				 ORDER BY DET.intNumeroParcela
			OPEN    c_BuscaCamposDetalheDividaAtiva
			FETCH   c_BuscaCamposDetalheDividaAtiva INTO @intDividaAtiva, @intNumeroPagina,
				@intNumeroInscricao, @dtmDataInscricao, @intNumeroLivroInscricao
			CLOSE c_BuscaCamposDetalheDividaAtiva
			DEALLOCATE c_BuscaCamposDetalheDividaAtiva

			EXECUTE (@strSql)

			SET @dblValorDesconto = 0
			IF @intDesconto <> 0
				SET @dblValorDesconto = ROUND((@dblValorTotal) * (@intDesconto / 100.00),2)

			SET @dblValorTotal = @dblValorTotal - @dblValorDesconto
			SET @dblValorParcelaReceita = ROUND((@dblValorTotal / @intParcelas),2)
			SET @dblValorDescontoParcela = ROUND((@dblValorDesconto / @intParcelas),2)
			
			-- Gravar as Parcelas Taxas
			EXECUTE sp_EfetuaCalculo @strPKIdReceita, @intComposicaoDaReceita ,21,
				@intParcelas, @dtmDataVencimento, @intIntervalo,0,0,
				@intCodUsuario
			-- Fim Gravar 

			-- Loop para gravar a parcela receita e Detalhe da Dívida Ativa
			SET @intContaPagina = 1
			SET @intIndiceParcelaReceita = 1
			SET @dblValorResto = 0
			SET @dblValorRestoDesconto = 0
			WHILE @intIndiceParcelaReceita <= @intParcelas
			BEGIN
				IF @intIndiceParcelaReceita = @intParcelas
				BEGIN
					SET @dblValorParcelaReceita = (@dblValorTotal - @dblValorResto)
					SET @dblValorDescontoParcela = (@dblValorDesconto - @dblValorRestoDesconto)
				END
				ELSE
				BEGIN
					SET @dblValorResto = (@dblValorResto + @dblValorParcelaReceita)
					SET @dblValorRestoDesconto = (@dblValorRestoDesconto + @dblValorDescontoParcela)
				END
			
				INSERT INTO tblParcelaReceita
					   (intLancamentoCalculo, intComposicaoDaReceita, intNumeroParcela, 
					    dtmDataVencimento, dblValorParcela, dblValorDesconto,
					    bytDividaAjuizada, bytSimulado, bytPrescrita, bytCancelada, 
					    bytAtiva, bytSuspensaoDeExigencia, dtmDtAtualizacao, lngCodUsr)
					   (SELECT MAX(PKId), @intComposicaoDaReceita, 
						   @intIndiceParcelaReceita, @dtmDataVencimento, 
						   @dblValorParcelaReceita, @dblValorDescontoParcela, 
						   0, 0, 0, 0, 1,0, GETDATE(), @intCodUsuario 
					      FROM tblLancamentoCalculo) 

				IF @intContaPagina > 60
				BEGIN
					SET @intNumeroPagina = @intNumeroPagina + 1
					SET @intContaPagina = 1
				END

				INSERT INTO tblDetalheDividaAtiva
					   (intDividaAtiva, strInscricaoCadastral, intExercicio, 
					    dtmVencimento, intNumeroParcela, dtmInscricao, 
					    intComposicaoReceita, intOcorrencia, dblValorOriginal, 
					    dblValorAtual, bytOrigem, bytDebitoGeradoManualmente,
	 				    bytSituacao, intNumeroLivroInscricao, 
					    intNumeroPaginaInscricao, intNumeroInscricao,
					    dtmDtAtualizacao, lngCodUsr ) VALUES 
					   (@intDividaAtiva, @strInscricaoCadastral, 
					    YEAR(GETDATE()), @dtmDataVencimento, 
					    @intIndiceParcelaReceita, @dtmDataInscricao, 
					    @intComposicaoDaReceita, @intOcorrencia, 
					    @dblValorParcelaReceita, @dblValorParcelaReceita, 
					    @bytOrigem,	0, 2, @intNumeroLivroInscricao, 
					    @intNumeroPagina, @intNumeroInscricao, GETDATE(),
		 			    @intCodUsuario) 
 
				SET @intContaPagina = @intContaPagina + 1
				SET @intNumeroInscricao = @intNumeroInscricao + 1
				SET @dtmDataVencimento = @dtmDataVencimento + @intIntervalo
				SET @intIndiceParcelaReceita = @intIndiceParcelaReceita + 1
			END 
			FETCH   c_ParcelaReceita INTO
				@dblValorTotal, @intComposicaoDaReceita, @intExercicio, 
				@dtmLancamento,	@bitUtilizacaoDebito, @bytOrigem
		END
		SET @bytResultado = 1
	END
	ELSE SET @bytResultado = 0
	
	CLOSE c_ParcelaReceita
	DEALLOCATE c_ParcelaReceita
	SELECT @bytResultado 
