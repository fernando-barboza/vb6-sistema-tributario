IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_EfetuaCalculoTaxa' AND TYPE = 'P')
   DROP PROCEDURE sp_EfetuaCalculoTaxa
GO

CREATE PROCEDURE sp_EfetuaCalculoTaxa(@dblValor   AS MONEY,
				      @intReceita AS INT,
				      @intFlag	  AS INT,
				      ---------------------------------------------------
				      -- Form's Com Data, intervalo e Nº de Parcelas
				      @intParcela AS INT = 0,
				      @dtmDtParcela AS DATETIME,
				      @intIntervalo AS INT,
                                      ---------------------------------------------------
				      -- Form's Com Data de vencimento Gravada em Tabela
				      @intTributo	  AS INT,
                                      ---------------------------------------------------
				      @gLngCodUsr AS INT)
AS

	DECLARE @intLancamentoCalculo AS INT,
		@intQuantidadeParcela AS INT,
		@i		      AS INT,
		@dtmDataVencimento    AS DATETIME,
		@dblParcelaDizma      AS MONEY,
		@dblResto	      AS DECIMAL(28,2),
 		@dblValorParcela AS DECIMAL(28,2)


	SET @intLancamentoCalculo = (SELECT MAX(PKId) 
			               FROM tblLancamentoCalculo)
	SET @dblValorParcela = 0
	SET @dblParcelaDizma = 0
	SET @dblResto = 0

	-- Caminhos Básicos   1º, 2º OU 3º

	IF @intFlag = 21 -- 1º Caminho
	BEGIN
		SET @intQuantidadeParcela = (@intParcela)
		SET @dtmDataVencimento = (@dtmDtParcela)
		SET @i = 1
		SET @dblParcelaDizma = @dblValor / @intQuantidadeParcela
		-- "For" Inserindo as Parcelas, quando as Mesmas forem definidas por Intervalo
		WHILE NOT @i > @intQuantidadeParcela
		BEGIN
			IF @i = @intQuantidadeParcela
				SET @dblValorParcela = @dblValor - @dblResto
			ELSE
			BEGIN
				SET @dblValorParcela = ROUND(@dblParcelaDizma,2)
				SET @dblResto = @dblResto + @dblValorParcela
			END
			IF @i <> 1 --Data Vencimento, vinda do Form + Intervalo	
				SET @dtmDataVencimento = DATEADD(DAY, @intIntervalo, @dtmDataVencimento)
				INSERT INTO tblParcelaTaxa SELECT
					@intLancamentoCalculo, @intReceita, @i, @dtmDataVencimento,
					@dblValorParcela, GETDATE(), @gLngCodUsr
				SET @i = @i + 1
			END	
		END
	ELSE
		IF @intFlag = 22  -- 2º Caminho
		BEGIN 
			SET @intQuantidadeParcela = (SELECT COUNT(VP.PKId) AS Quantidade 
						       FROM tblVencimentosDasParcelas VP,
							    tblVencimentos VC
						      WHERE VC.PKId = VP.intVencimento
							AND VC.intTributo = @intTributo)
			SET @dblParcelaDizma = @dblValor / @intQuantidadeParcela
			CREATE TABLE #t_EfetuaCalculoTaxa
				(i			INT,
				 dtmDataVencimento DATETIME)
			INSERT INTO #t_EfetuaCalculoTaxa
				SELECT VP.intNumero , VP.dtmDataDaParcela
				   FROM tblVencimentosDasParcelas VP,
					tblVencimentos VC
				  WHERE VC.PKId = VP.intVencimento 
				    AND VC.intTributo = @intTributo
				ORDER BY VP.intNumero
			-- Cursor
			-- Inserir Quando as Parcela ja estiverem na Tabela tblVencimentosDasParcelas
			DECLARE	c_EfetuaCalculoTaxa CURSOR FOR
				SELECT i, dtmDataVencimento  FROM #t_EfetuaCalculoTaxa
			OPEN	c_EfetuaCalculoTaxa
			FETCH	c_EfetuaCalculoTaxa INTO
				@i, @dtmDataVencimento
			WHILE @@FETCH_STATUS = 0
			BEGIN
				IF @i = @intQuantidadeParcela
					SET @dblValorParcela = @dblValor - @dblResto
				ELSE
				BEGIN
					SET @dblValorParcela = ROUND(@dblParcelaDizma,2)
					SET @dblResto = @dblResto + @dblValorParcela
				END
	
				INSERT INTO tblParcelaTaxa SELECT
					@intLancamentoCalculo, @intReceita, @i, @dtmDataVencimento,
					@dblValorParcela, GETDATE(), @gLngCodUsr

				FETCH c_EfetuaCalculoTaxa INTO
				      @i, @dtmDataVencimento		
			END
			-- Fim Cursor
		END
		ELSE -- 3º Caminho  @intFlag = 23      
			INSERT INTO tblParcelaTaxa SELECT  
				@intLancamentoCalculo, @intReceita, @intParcela, @dtmDtParcela,
				@dblValor, GETDATE(), @gLngCodUsr



