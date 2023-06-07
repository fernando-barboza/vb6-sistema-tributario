IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoLancamentoReceitas' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoLancamentoReceitas
GO

/*'15,9,30,7,71'
sp_CalculoLancamentoReceitas 1, 18 , 3.2,'11,15',
' SELECT A.intContribuinte, A.strInscricaoAnterior FROM tblImobiliario AS A, tblContribuinte AS CO  WHERE  CO.PKId = A.intContribuinte  AND A.strInscricaoAnterior = "01010010010001"',
2001,7, NULL, '2001/01/01 00:00:00','2001/01/01 00:00:00',1,6,10,01,8, 1,45,0,0,50.0000,
'?50.00,40,10, @dblValor OUTPUT'
*/
CREATE PROCEDURE sp_CalculoLancamentoReceitas(@intFlag	  	    AS INT,
					      @dblValorAparcelar    AS MONEY,
					      @dblValorNaoParcelado AS MONEY,
					      @strPKId		    AS NVARCHAR(4000),
					      @strContribuinte	    AS NVARCHAR(1000),--
					      @intExercicio	    AS INT,
					      @intComposicaoReceita AS INT,
					      @intMensagem	    AS INT,
					      @dtmLancamento	    AS DATETIME,
					      @dtmVencimento	    AS DATETIME,
					      @intParcelaInicial    AS INT ,
					      @intParcelaFinal      AS INT ,		
					      @intIntervalo 	    AS INT,
					      @bitUtilizacaoDebito  AS TINYINT,
					      @intOcorrencia	    AS INT,
					      @bytOrigem	    AS TINYINT,
					      @gLngCodUsr   	    AS INT,
					      @PKIdImobilEconomico  INT = 0,
					      @dblValorDesconto	    MONEY = NULL,
					      @dblValorAliquota	    MONEY = NULL,
					      @strParametro	    NVARCHAR(100) = NULL)

AS
	DECLARE @PKIdContribuinte       AS INT,
		@strInscricaoCadastral  AS NVARCHAR(50),
		@PKIdAtividadePrincipal AS INT,
		@strAtividadePrincipal  AS NVARCHAR(100),
		@strSequencia	        AS NVARCHAR(50),
		@dblValorDiferenca	AS MONEY,
		@dblAliquota		AS MONEY

	--CONFIGURAÇÔES DE PARAMETROS E ALIQUOTAS--------------------------------------------
	IF @dblValorAliquota = (-1)
		SET @dblAliquota = NULL
	ELSE
		IF SUBSTRING(@strParametro,1,1) = '?'
		BEGIN
			SET @strParametro = SUBSTRING(@strParametro,2,LEN(@strParametro))
			SET @dblAliquota = @dblValorAliquota
			SET @dblValorAliquota = -1
		END
		ELSE
			SET @dblAliquota = @dblValorAliquota
	------------------------------------------------------------------------------------

	EXECUTE('DECLARE c_CalculoLancamentoReceitas CURSOR FOR ' + @strContribuinte )
 
	OPEN	c_CalculoLancamentoReceitas
	FETCH	c_CalculoLancamentoReceitas INTO
		@PKIdContribuinte, @strInscricaoCadastral
	WHILE @@FETCH_STATUS = 0
	BEGIN

		--INICIALIZAÇÕES DE VARIÁVEIS----------------------------------------------
		SET @strSequencia = (SELECT ISNULL(MAX(strSequencia),0) + 1 AS Maximo 	 --
				       FROM tblLancamentoCalculo			 --
				      WHERE intComposicaoReceita = @intComposicaoReceita --
					AND intContribuinte = @PKIdContribuinte 	 --
					AND intExercicio = @intExercicio )		 --
		--FIM DA INICIALIZAÇÃO-----------------------------------------------------

		--EFETUA INSERÇÕES NA TABELA tblLancamentoCalculo------------------------------
		INSERT INTO tblLancamentoCalculo (intExercicio, intContribuinte, 	     --
			    intComposicaoReceita, intMensagem, strInscricaoCadastral,  	     --
			    dtmLancamento, dtmVencimento, intNumeroDeParcelas, 		     --
			    intIntervaloEntreParcelas,  bitUtilizacaoDebito, intOcorrencia,  --
			    bytOrigem, dblAliquota, strSequencia, dtmDtAtualizacao,lngCodUsr)--
			    VALUES ( @intExercicio, @PKIdContribuinte, @intComposicaoReceita,--
			    @intMensagem, @strInscricaoCadastral, @dtmLancamento,	     --
			    @dtmVencimento, @intParcelaFinal - @intParcelaInicial + 1, 	     --
			    @intIntervalo, @bitUtilizacaoDebito, @intOcorrencia, @bytOrigem, --
		            @dblAliquota, @strSequencia, GETDATE(), @gLngCodUsr)	     --
		--FIM DAS INSERÇÕES NA TABELA tblLancamentoCalculo-----------------------------					

		--SE NÃO FOI CALCULADO ANTERIORMENTE O VALOR A PARCELAR, ENTÃO CALCULE AGORA-------
		IF (@dblValorNaoParcelado < (0.000))						 --
			EXECUTE sp_CalculoParaUsuario @strInscricaoCadastral, @strPKId,		 --
						      @bytOrigem, @intComposicaoReceita,	 --
						      @dblValorNaoParcelado OUTPUT,		 --
						      @dblValorAparcelar OUTPUT, 		 --	
						      @PKIdImobilEconomico , @strParametro	 --
		-----------------------------------------------------------------------------------
		SET @dblValorDiferenca = (@dblValorNaoParcelado + @dblValorAparcelar)
		--EFETUA INSERÇÕES NA TABELA tblParcelaReceita--------------------------------------------
		EXECUTE sp_CalculoParcelaReceita  @intFlag, @intComposicaoReceita ,@dblValorAparcelar , --
						  @dblValorNaoParcelado, @intParcelaInicial, 		--
						  @intParcelaFinal, @intExercicio, @intIntervalo, 	--
						  @dtmVencimento, @gLngCodUsr, @dblValorDesconto	--
		--FIM DAS INSERÇÕES NA TABELA tblParcelaReceita-------------------------------------------				

		--EFETUA INSERÇÕES NA TABELA tblParcelaTaxa---------------------------------------------
		EXECUTE sp_CalculoParcelaTaxa @strInscricaoCadastral,@intFlag,@strPKId, @bytOrigem,   --
					      @intExercicio, @intComposicaoReceita,@dblValorDiferenca,--
					      @intParcelaInicial, @intParcelaFinal, @dtmVencimento,   --
					      @intIntervalo, @gLngCodUsr, @PKIdImobilEconomico,       --		
					      @dblValorDesconto, @dblValorAliquota, @strParametro     --
		--FIM DAS INSERÇÕES NA TABELA tblParcelaTaxa--------------------------------------------				
		FETCH c_CalculoLancamentoReceitas INTO
		      @PKIdContribuinte, @strInscricaoCadastral
	END
	CLOSE c_CalculoLancamentoReceitas
	DEALLOCATE c_CalculoLancamentoReceitas



