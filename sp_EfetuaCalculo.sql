IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_EfetuaCalculo' AND TYPE = 'P')
   DROP PROCEDURE sp_EfetuaCalculo
GO

       CREATE PROCEDURE sp_EfetuaCalculo(@strPKId 		AS NVARCHAR(100),
/*------------------------------------*/ @intComposicaoReceita 	AS INT,
/* Valores Do @intFlag                */ @intFlag		AS INT,
/*				      */ -------------------------------------------------------
/* ---- SEM TAXA ----   	      */ -- (31) Form's Com Data, intervalo e Nº de Parcelas 
/* 11 -- Valor Calculado + Indexador  */ @intParcela 		AS INT = 0,      
/* 12 -- Item 11 + Desconto           */ @dtmDtParcela 		AS DATETIME,
/*                                    */ @intIntervalo 		AS INT,
/* 		                      */ -------------------------------------------------------
/* ---- COM TAXA ----                 */ -- (32) Form's Com Data de vencimento Gravada em Tabela 
/* 21 ->Data,Intervalo,NºParcelas     */ @intTributo	  	AS INT,  
/* 22 ->Datas da tabela tblVencimento */ -------------------------------------------------------
/* 23 ->Transferência Divida Ativa    */ -- (2) Form's Com Desconto 
/*------------------------------------*/ @Desconto 		NUMERIC = 100 OUTPUT,
				 	 -------------------------------------------------------
				 	 @gLngCodUsr 		AS INT)

AS
	CREATE TABLE #t_EfetuaCalculo
		(PKId		INT)

	DECLARE @PKId		    AS INT,
		@ContaParcelas	    AS INT,
		@blnUsaFaixaDeValor AS TINYINT,
		@blnECalculada 	    AS TINYINT,
		@dblValor	    AS MONEY,
		@dblIndexador	    AS MONEY,
		@dblValorCalculado  AS MONEY,
		@dblExecFormula	    AS MONEY,
		@dblValorProcedure  AS MONEY,
		@strExecFormula	    AS NVARCHAR(1000),
		@intPosition	    AS INT,
		@intCode	    AS INT,
		@blnNomeInvalido    AS TINYINT,
		@IntImposto	    AS TINYINT,
		@dblDesconto	    AS MONEY	

	SET @dblValorCalculado = 0
	SET @dblDesconto = 0

--Cálcula Indexador
	SET @dblIndexador = (Select IE.dblValor
			       FROM tblIndiceEconomico  IE, tblIndexadorEconomico E 
			      WHERE IE.intIndexador = E.PKId
				AND IE.dtmData in (SELECT MAX(IE.dtmData) 
						     FROM tblIndiceEconomico IE))
--Fim do Cálculo Indexador

	EXECUTE('INSERT INTO #t_EfetuaCalculo
			SELECT PKid FROM tblReceita
			WHERE PKId IN('+@strPKId+')')

	DECLARE	c_EfetuaCalculo CURSOR FOR
		SELECT PKId FROM #t_EfetuaCalculo
	OPEN	c_EfetuaCalculo
	FETCH	c_EfetuaCalculo INTO
		@PKId
	WHILE @@FETCH_STATUS = 0
	BEGIN
		SET @blnUsaFaixaDeValor = (SELECT blnUsaFaixaDeValor FROM tblReceita WHERE PKId = @PKId)
		SET @blnECalculada = (SELECT blnECalculada FROM tblReceita WHERE PKId = @PKId)
		SET @dblValor = (SELECT dblValor FROM tblReceita WHERE PKId = @PKId)
		SET @intImposto = (SELECT bytTipo FROM	tblReceita WHERE PKId = @PKId)

		IF @blnECalculada = 1 
		BEGIN
			--Referente à Formula
			SET @strExecFormula = (SELECT 'sp_' + FC.strNome FROM tblReceita RC, tblFormulaDeCalculo FC
						  WHERE FC.PKId =  RC.intFormuladeCalculo
						    AND RC.PKId = @PKId)
			SET @intPosition = 1
			WHILE @intPosition <= DATALENGTH(@strExecFormula)
			BEGIN
				SET @intCode = (SELECT ASCII(SUBSTRING(@strExecFormula, @intPosition, 1)))
				IF @intCode = 32
					SET @blnNomeInvalido = 1
	 			SET @intPosition = @intPosition + 1
			END
		
			IF @blnNomeInvalido = 0
				EXECUTE('SET @dblExecFormula = ('+ @strExecFormula +')')
			ELSE
				SET @dblExecFormula = 0
			--Fim da Fórmula
			IF @blnUsaFaixaDeValor = 1 
			BEGIN
				SET @dblValor = (SELECT F.dblValor FROM	tblComposicaoDaReceita A,
						tblValorCompoRec B, tblReceita C, tblFaixaDeValor D,		
						tblValorDaFaixa E, tblTabelaDeValor F 
						WHERE B.intReceita = C.PKId
						  AND A.PKId = B.intComposicaoDaReceita
						  AND D.PKId = E.intFaixaDeValores
						  AND C.intFaixaDeValor = D.PKId
						  AND F.PKId = E.intTabelaDeValores
						  AND A.PKId = @intComposicaoReceita 
						  AND C.PKId = @PKId)
				-- Com Desconto
				IF @intFlag = 12
					IF @intImposto = 2 
						SET @dblDesconto = @dblDesconto + (@dblValor * @Desconto / 100)
				-- Fim Desconto
				IF @dblIndexador > 0
				BEGIN 
					-- Sem Suporte A TAXA
					IF ROUND(@intFlag/10, 0) = 1
						SET @dblValorCalculado = (@dblValorCalculado + @dblExecFormula) + (@dblValor *@dblIndexador)
					-- Fim sem Suporte TAXA
					-- Com Suporte A TAXA
					IF ROUND(@intFlag/10, 0) = 2
					BEGIN	
						SET @dblValorProcedure = @dblValor * @dblIndexador
						EXECUTE sp_EfetuaCalculoTaxa @dblValorProcedure, @PKId,@intFlag,
						        @intParcela, @dtmDtParcela, @intIntervalo, @intTributo, 
							@gLngCodUsr
					END
					-- Fim Com Suporte TAXA
				END
			END
			ELSE
			BEGIN
				-- Com Desconto
				IF @intFlag = 12
					IF @intImposto = 2
						SET @dblDesconto = @dblDesconto + (@dblExecFormula * @Desconto / 100)
				-- Fim Desconto
				-- Sem Suporte A TAXA
				IF ROUND(@intFlag/10, 0) = 1
					SET @dblValorCalculado = @dblValorCalculado + @dblExecFormula
				-- Fim Sem Suporte TAXA
				-- Com Suporte A TAXA
				IF ROUND(@intFlag/10, 0) = 2
					EXECUTE sp_EfetuaCalculoTaxa @dblexecformula, @PKId, @intFlag, 
						@intParcela, @dtmDtParcela, @intIntervalo, @intTributo, 
						@gLngCodUsr
				-- Fim Com Suporte TAXA
			END
		END
		ELSE
		BEGIN
			-- Com Desconto
			IF @intFlag = 12
				IF @intImposto = 2
					SET @dblDesconto = @dblDesconto + (@dblValor * @Desconto / 100)
			-- Fim Desconto
			-- Sem Suporte A TAXA
			IF ROUND(@intFlag/10, 0) = 1
				SET @dblValorCalculado = @dblValorCalculado + @dblValor
			-- Fim Sem Suporte TAXA
			-- Com Suporte A TAXA
			IF ROUND(@intFlag/10, 0) = 2
				EXECUTE sp_EfetuaCalculoTaxa @dblValor, @PKId, @intFlag, 
					@intParcela, @dtmDtParcela, @intIntervalo, @intTributo, 
					@gLngCodUsr
		END
		FETCH c_EfetuaCalculo INTO
		      @PKId
	END
	CLOSE c_EfetuaCalculo
	DEALLOCATE c_EfetuaCalculo
	IF @intFlag = 11
		SELECT ISNULL(@dblValorCalculado,0) AS dblValorCalculado ,
		       ISNULL(@dblIndexador,0) AS dblIndexador
	ELSE
		IF @intFlag = 12
			SELECT ISNULL(@dblValorCalculado,0) AS dblValorCalculado ,
			       ISNULL(@dblIndexador,0) AS dblIndexador,
		       	       ISNULL(@dblDesconto,0) AS dblDesconto


		