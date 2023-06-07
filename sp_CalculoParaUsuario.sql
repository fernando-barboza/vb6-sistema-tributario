IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoParaUsuario' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoParaUsuario
GO
-- sp_CalculoParaUsuario '000004','10,15',2,6
CREATE PROCEDURE sp_CalculoParaUsuario(@strInscricaoCadastral AS NVARCHAR(100),
				       @strPKId		      AS NVARCHAR(4000),
				       @bytUtilizacao	      AS TINYINT,
				       @intComposicaoReceita  AS INT,
				       @dblValorNaoParcelado  MONEY = 0 OUTPUT,
				       @dblValorAparcelar     MONEY = 0 OUTPUT,
				       @PKIdImobilEconomico   INT = 0,
				       @strParametro	      NVARCHAR(100) = ' @dblValor OUTPUT')
AS
	CREATE TABLE #t_CalculoLancamento
		     (PKId 	      INT)

	EXECUTE('INSERT INTO #t_CalculoLancamento
			SELECT PKid FROM tblReceita
			WHERE PKId IN('+@strPKId+')')

	--VARIÁVEIS A SEREM USADAS----------------------
	DECLARE @PKId		      AS INT,  	      --
		@blnUsaFaixaDeValor   AS BIT,	      --	
		@blnECalculada 	      AS BIT,	      --
		@blnParcelar	      AS BIT,	      --
		@dblValor	      AS MONEY,	      --
		@dblIndexador	      AS MONEY,	      --	
		@dblExecFormula	      AS MONEY,	      --
		@strParametrosAUX     AS NVARCHAR(100)--
	--FIM DA DECLARAÇÃO ----------------------------

	--INICIALIZAÇÕES DE VARIÁVEIS-------------------------------------------------
	SET @dblValorNaoParcelado = 0						    --			
	SET @dblValorAparcelar = 0						    --
	SET @strParametrosAUX = @strParametro					    --
	SET @dblIndexador = (SELECT IE.dblValor					    --
			       FROM tblIndiceEconomico  IE, tblIndexadorEconomico E -- 
			      WHERE IE.intIndexador = E.PKId			    --
				AND IE.dtmData IN (SELECT MAX(IE.dtmData) 	    --
						     FROM tblIndiceEconomico IE))   --
	IF @bytUtilizacao = 2							    --
		SET @PKIdImobilEconomico = (SELECT PKId 			    --
					      FROM tblEconomico 		    --
					     WHERE strInscricaoCadastral = @strInscricaoCadastral)
	--FIM DA INICIALIZAÇÃO--------------------------------------------------------		

	DECLARE	c_CalculoLancamento CURSOR FOR
		SELECT PKId FROM #t_CalculoLancamento
	OPEN	c_CalculoLancamento
	FETCH	c_CalculoLancamento INTO
		@PKId
	WHILE @@FETCH_STATUS = 0
	BEGIN
		SELECT @blnUsaFaixaDeValor = blnUsaFaixaDeValor, @blnECalculada = blnECalculada,
		       @dblValor = dblValor, @blnParcelar = blnParcelar
		  FROM tblReceita 
		 WHERE PKId = @PKId  
		IF @blnParcelar = 0
			IF @blnECalculada = 1
			BEGIN
				EXECUTE sp_ParametroReceitas @PKId, @PKIdImobilEconomico, @strParametro OUTPUT
				EXECUTE sp_CalculoFormulaExecutada @PKId, @dblExecFormula OUTPUT, @strParametro
				SET @strParametro = @strParametrosAUX
				IF @blnUsaFaixaDeValor = 1 
				BEGIN
					SET @dblValor = (SELECT F.dblValor 
							   FROM	tblComposicaoDaReceita A,
								tblValorCompoRec B, tblReceita C, 
								tblFaixaDeValor D, tblValorDaFaixa E, 
								tblTabelaDeValor F 
							  WHERE B.intReceita = C.PKId
							    AND A.PKId = B.intComposicaoDaReceita
							    AND D.PKId = E.intFaixaDeValores
							    AND C.intFaixaDeValor = D.PKId
							    AND F.PKId = E.intTabelaDeValores
							    AND A.PKId = @intComposicaoReceita 
							    AND C.PKId = @PKId)
					IF @dblIndexador > 0
					BEGIN
						SET @dblValor = ISNULL(@dblExecFormula,0) + 
								(ISNULL(@dblValor,0) *ISNULL(@dblIndexador,0))
						EXECUTE sp_VerificaIsencao @bytUtilizacao, @strInscricaoCadastral, 
									   @intComposicaoReceita, @PKId, 
									   @dblValor OUTPUT
						SET @dblValorAparcelar = @dblValorAparcelar + @dblValor
					END
				END
				ELSE
				BEGIN
					EXECUTE sp_VerificaIsencao @bytUtilizacao, @strInscricaoCadastral, 
								   @intComposicaoReceita, @PKId, 
								   @dblExecFormula OUTPUT
					SET @dblValorAparcelar = @dblValorAparcelar + ISNULL(@dblExecFormula,0)
				END
			END
			ELSE
			BEGIN
				EXECUTE sp_VerificaIsencao @bytUtilizacao, @strInscricaoCadastral, 
							   @intComposicaoReceita, @PKId, 
							   @dblValor OUTPUT
				SET @dblValorAparcelar = @dblValorAparcelar + ISNULL(@dblValor,0)
			END
		ELSE
		BEGIN
			EXECUTE sp_VerificaIsencao @bytUtilizacao, @strInscricaoCadastral, 
						   @intComposicaoReceita, @PKId, 
						   @dblValor OUTPUT
			SET @dblValorNaoParcelado = @dblValorNaoParcelado + ISNULL(@dblValor,0)
		END
		FETCH c_CalculoLancamento INTO
		      @PKId
	END
	CLOSE c_CalculoLancamento
	DEALLOCATE c_CalculoLancamento
	IF @intComposicaoReceita != 16 AND @intComposicaoReceita != 3 AND @intComposicaoReceita != 4
		SELECT ISNULL(@dblValorNaoParcelado,0) AS dblValorNaoParcelado,
		       ISNULL(@dblValorAparcelar,0)    AS dblValorAparcelar



