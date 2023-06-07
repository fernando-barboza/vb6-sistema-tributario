IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoParcelaReceita' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoParcelaReceita
GO

CREATE PROCEDURE sp_CalculoParcelaReceita(@intFlagReceita   	AS INT,
				      	  @intComposicaoReceita AS INT,
 				          @dblValorAparcelar    AS MONEY,
				          @dblValorNaoParcelado AS MONEY,
				     	  @intParcelaInicial	AS INT,
				     	  @intParcelaFinal	AS INT,
					  @intExercicio		AS INT,			
				     	  @intIntervalo 	AS INT,
				      	  @dtmVencimento	AS DATETIME,
				    	  @gLngCodUsr   	AS INT,
					  @dblValorDesconto     MONEY = 0)					  
AS
	DECLARE @PKIdLancamentoCalculo AS INT,
		@intQuantidadeParcela  AS INT,
		@i		       AS INT,
		@dtmDataVencimento     AS DATETIME,
		@dblParcelaDizma       AS MONEY,
		@dblResto	       AS DECIMAL(28,2),
 		@dblValorParcela       AS DECIMAL(28,2)
		
	SET @PKIdLancamentoCalculo = (SELECT MAX(PKId) FROM tblLancamentoCalculo)
	SET @dblValorParcela = 0
	SET @dblParcelaDizma = 0
	SET @dblResto = 0
	
	IF @intFlagReceita > 2
	BEGIN 	
		SELECT @i = VP.intNumero, @dtmDataVencimento = VP.dtmDataDaParcela
   	          FROM tblVencimentosDasParcelas VP, tblVencimentos VC
		 WHERE VC.PKId = VP.intVencimento
		   AND VP.intNumero BETWEEN @intParcelaInicial AND @intParcelaFinal
		   AND VC.intTributo = @intFlagReceita
		   AND VP.intExercicio = @intExercicio
		   AND VP.intNumero = 0
		IF @i = 0 --SE EXISTE PARCELA ZERO
			INSERT INTO tblParcelaReceita		
	 			   (intLancamentoCalculo, intComposicaoDaReceita, intNumeroParcela, 
			 	    dtmDataVencimento, dblValorParcela, bytDividaAjuizada, bytSimulado, 
				    bytPrescrita, bytCancelada, bytAtiva, bytSuspensaoDeExigencia, 
				    dtmDtAtualizacao, lngCodUsr) VALUES (@PKIdLancamentoCalculo,
				    @intComposicaoReceita, @i, @dtmDataVencimento , (@dblValorAparcelar
				    - (@dblValorAparcelar * @dblValorDesconto/100)+ @dblValorNaoParcelado),
				     0,0,0,0,0,0,GETDATE(), @gLngCodUsr)

		SET @intQuantidadeParcela = (SELECT COUNT(VP.PKId) AS Quantidade 
					       FROM tblVencimentosDasParcelas VP,
						    tblVencimentos VC
					      WHERE VC.PKId = VP.intVencimento
						AND VP.intNumero BETWEEN @intParcelaInicial AND @intParcelaFinal
						AND VC.intTributo = @intFlagReceita
						AND VP.intExercicio = @intExercicio
						AND VP.intNumero != 0)

		SET @dblParcelaDizma = @dblValorAparcelar / @intQuantidadeParcela
		CREATE TABLE #t_EfetuaCalculoTaxa
			(i			INT,
			 dtmDataVencimento DATETIME)
		INSERT INTO #t_EfetuaCalculoTaxa
			SELECT VP.intNumero , VP.dtmDataDaParcela
			   FROM tblVencimentosDasParcelas VP,
				tblVencimentos VC
			  WHERE VC.PKId = VP.intVencimento 
			    AND VP.intNumero BETWEEN @intParcelaInicial AND @intParcelaFinal
			    AND VC.intTributo = @intFlagReceita
			    AND VP.intExercicio = @intExercicio
			    AND VP.intNumero != 0
			ORDER BY VP.intNumero
		-- INSERIR QUANDO AS PARCELAS JA ESTIVEREM 
		-- NA TABELA tblVencimentosDasParcelas
		DECLARE	c_CalculoParcelaReceita CURSOR FOR
			SELECT i, dtmDataVencimento  FROM #t_EfetuaCalculoTaxa
		OPEN	c_CalculoParcelaReceita
		FETCH	c_CalculoParcelaReceita INTO
			@i, @dtmDataVencimento
		WHILE @@FETCH_STATUS = 0
		BEGIN
			IF @i = @intParcelaFinal
				SET @dblValorParcela = @dblValorAparcelar - @dblResto
			ELSE
			BEGIN
				SET @dblValorParcela = ROUND(@dblParcelaDizma,2)
				SET @dblResto = @dblResto + @dblValorParcela
			END
			INSERT INTO tblParcelaReceita		
	 			   (intLancamentoCalculo, intComposicaoDaReceita, intNumeroParcela, 
			 	    dtmDataVencimento, dblValorParcela, bytDividaAjuizada, bytSimulado, 
				    bytPrescrita, bytCancelada, bytAtiva, bytSuspensaoDeExigencia, 
				    dtmDtAtualizacao, lngCodUsr) VALUES (@PKIdLancamentoCalculo,
				    @intComposicaoReceita, @i, @dtmDataVencimento , (@dblValorNaoParcelado 
				    + @dblValorParcela), 0,0,0,0,0,0,GETDATE(), @gLngCodUsr)
			FETCH c_CalculoParcelaReceita INTO
			      @i, @dtmDataVencimento		
		END
	END 
	ELSE IF @intFlagReceita = 1
	BEGIN
		SET @intQuantidadeParcela = (@intParcelaFinal - @intParcelaInicial + 1)
		SET @dtmDataVencimento = (@dtmVencimento)
		SET @i = @intParcelaInicial
		SET @dblParcelaDizma = @dblValorAparcelar / @intQuantidadeParcela
		-- "LOOP" INSERINDO AS PARCELAS, QUANDO AS MESMAS FOREM DEFINIDAS POR INTERVALO
		WHILE NOT @i > @intParcelaFinal
		BEGIN
			IF @i = @intParcelaFinal
				SET @dblValorParcela = @dblValorAparcelar - @dblResto
			ELSE
			BEGIN
				SET @dblValorParcela = ROUND(@dblParcelaDizma,2)
				SET @dblResto = @dblResto + @dblValorParcela
			END
			IF @i <> @intParcelaInicial --Data Vencimento, vinda do Form + Intervalo	
				SET @dtmDataVencimento = DATEADD(DAY, @intIntervalo, @dtmDataVencimento)
			INSERT INTO tblParcelaReceita		
	 			   (intLancamentoCalculo, intComposicaoDaReceita, intNumeroParcela, 
			 	    dtmDataVencimento, dblValorParcela, bytDividaAjuizada, bytSimulado, 
				    bytPrescrita, bytCancelada, bytAtiva, bytSuspensaoDeExigencia, 
				    dtmDtAtualizacao, lngCodUsr) VALUES (@PKIdLancamentoCalculo,
				    @intComposicaoReceita, @i, @dtmDataVencimento , (@dblValorNaoParcelado 
				    + @dblValorParcela), 0,0,0,0,0,0,GETDATE(), @gLngCodUsr)					
			SET @i = @i + 1
		END	
	END

