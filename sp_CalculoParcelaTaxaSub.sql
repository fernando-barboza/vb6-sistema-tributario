IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoParcelaTaxaSub' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoParcelaTaxaSub
GO

CREATE PROCEDURE sp_CalculoParcelaTaxaSub(@dblValor   	   	AS MONEY,
				      	  @intReceita 	    	AS INT,
				    	  @intFlagReceita	AS INT,
				          @intParcelaInicial    AS INT ,
					  @intParcelaFinal      AS INT ,	
				  	  @dtmDtParcela 	AS DATETIME,
					  @intExercicio		AS INT,
				   	  @intIntervalo 	AS INT,
				      	  @gLngCodUsr 		AS INT,
					  @dblValorDesconto	MONEY = 0)
AS

	DECLARE @intLancamentoCalculo AS INT,
		@intQuantidadeParcela AS INT,
		@i		      AS INT,
		@dtmDataVencimento    AS DATETIME,
		@dblParcelaDizma      AS MONEY,
		@dblResto	      AS DECIMAL(28,2),
 		@dblValorParcela      AS DECIMAL(28,2),
		@blnParcelar	      AS BIT


	SET @intLancamentoCalculo = (SELECT MAX(PKId) 
			               FROM tblLancamentoCalculo)
	SET @dblValorParcela = 0
	SET @dblParcelaDizma = 0
	SET @dblResto = 0
	SET @blnParcelar = (SELECT blnParcelar FROM tblReceita WHERE PKId = @intReceita)
	IF @intFlagReceita = 1 
	BEGIN
		SET @intQuantidadeParcela = (@intParcelaFinal - @intParcelaInicial + 1)
		IF (@blnParcelar = 1)
			SET @dblValor = @dblValor * @intQuantidadeParcela
		SET @dtmDataVencimento = (@dtmDtParcela)
		SET @i = @intParcelaInicial
		SET @dblParcelaDizma = @dblValor / @intQuantidadeParcela
		-- "LOOP" INSERINDO AS PARCELAS, QUANDO 
		-- AS MESMAS FOREM DEFINIDAS POR INTERVALO
		WHILE NOT @i > @intParcelaFinal
		BEGIN
			IF @i = @intParcelaFinal
				SET @dblValorParcela = @dblValor - @dblResto
			ELSE
			BEGIN
				SET @dblValorParcela = ROUND(@dblParcelaDizma,2)
				SET @dblResto = @dblResto + @dblValorParcela
			END
			IF @i <> @intParcelaInicial --Data Vencimento, vinda do Form + Intervalo	
				SET @dtmDataVencimento = DATEADD(DAY, @intIntervalo, @dtmDataVencimento)
				INSERT INTO tblParcelaTaxa SELECT
					@intLancamentoCalculo, @intReceita, @i, @dtmDataVencimento,
					@dblValorParcela, GETDATE(), @gLngCodUsr
				SET @i = @i + 1
			END	
		END
	ELSE  -- 2º Caminho
	BEGIN 
		SELECT @i = VP.intNumero, @dtmDataVencimento = VP.dtmDataDaParcela
	          FROM tblVencimentosDasParcelas VP, tblVencimentos VC
		 WHERE VC.PKId = VP.intVencimento
		   AND VC.intTributo = @intFlagReceita
		   AND VP.intNumero BETWEEN @intParcelaInicial AND @intParcelaFinal
		   AND YEAR(VP.dtmDataDaParcela) = @intExercicio
		   AND VP.intNumero = 0
		IF @i = 0 --SE EXISTE PARCELA ZERO
			IF (@blnParcelar = 0) -- A TAXA É DIVISÍVEL
				INSERT INTO tblParcelaTaxa SELECT
					@intLancamentoCalculo, @intReceita, @i, @dtmDataVencimento,
					(@dblValor - (@dblValor * @dblValorDesconto/100)), GETDATE(), 
					@gLngCodUsr
			ELSE
				INSERT INTO tblParcelaTaxa SELECT
					@intLancamentoCalculo, @intReceita, @i, @dtmDataVencimento,
					@dblValor, GETDATE(), @gLngCodUsr				
		SET @intQuantidadeParcela = (SELECT COUNT(VP.PKId) AS Quantidade 
					       FROM tblVencimentosDasParcelas VP,
						    tblVencimentos VC
					      WHERE VC.PKId = VP.intVencimento
						AND VC.intTributo = @intFlagReceita
						AND VP.intNumero BETWEEN @intParcelaInicial AND @intParcelaFinal
						AND VP.intExercicio = @intExercicio
						AND VP.intNumero != 0)
		IF (@blnParcelar = 1)
			SET @dblValor = @dblValor * @intQuantidadeParcela
		SET @dblParcelaDizma = @dblValor / @intQuantidadeParcela

		CREATE TABLE #t_EfetuaCalculoTaxa
			(i			INT,
			 dtmDataVencimento DATETIME)
		INSERT INTO #t_EfetuaCalculoTaxa
			SELECT VP.intNumero , VP.dtmDataDaParcela
			   FROM tblVencimentosDasParcelas VP,
				tblVencimentos VC
			  WHERE VC.PKId = VP.intVencimento 
			    AND VC.intTributo = @intFlagReceita
			    AND VP.intNumero BETWEEN @intParcelaInicial AND @intParcelaFinal
			    AND VP.intExercicio = @intExercicio
			    AND VP.intNumero != 0
			ORDER BY VP.intNumero
		-- INSERIR QUANDO AS PARCELAS JA ESTIVEREM 
		-- NA TABELA tblVencimentosDasParcelas
		DECLARE	c_EfetuaCalculoTaxa CURSOR FOR
			SELECT i, dtmDataVencimento  FROM #t_EfetuaCalculoTaxa
		OPEN	c_EfetuaCalculoTaxa
		FETCH	c_EfetuaCalculoTaxa INTO
			@i, @dtmDataVencimento
		WHILE @@FETCH_STATUS = 0
		BEGIN
			IF @i = @intParcelaFinal
				SET @dblValorParcela = @dblValor - @dblResto
			ELSE
			BEGIN
				SET @dblValorParcela = ROUND(@dblParcelaDizma,2)
				SET @dblResto = @dblResto + @dblValorParcela
			END

			INSERT INTO tblParcelaTaxa SELECT
				@intLancamentoCalculo, @intReceita, @i, @dtmDataVencimento,
				@dblValorParcela , GETDATE(), @gLngCodUsr
				FETCH c_EfetuaCalculoTaxa INTO
			      @i, @dtmDataVencimento		
		END
	END
