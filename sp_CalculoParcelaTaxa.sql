IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoParcelaTaxa' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoParcelaTaxa
GO
-- sp_CalculoParcelaTaxa 1,3.2000,'15,7,30',2000,3,10,'2000-01-01 00:00:00.000',30,45

CREATE PROCEDURE sp_CalculoParcelaTaxa(@strInscricaoCadastral AS NVARCHAR(100),
				       @intFlag		      AS INT,
				       @strPKId		      AS NVARCHAR(4000),
				       @bytUtilizacao	      AS TINYINT,
				       @intExercicio	      AS INT,	
				       @intComposicaoReceita  AS INT,
				       @dblValorDiferenca     AS MONEY,	 
				       @intParcelaInicial     AS INT ,
				       @intParcelaFinal       AS INT ,	
				       @dtmDtParcela 	      AS DATETIME,
				       @intIntervalo 	      AS INT,
				       @gLngCodUsr 	      AS INT,
				       @PKIdImobilEconomico   INT = 0,
				       @dblValorDesconto      MONEY = NULL,
				       @dblValorAliquota      MONEY = -1,
				       @strParametro	      NVARCHAR(100))
AS
	CREATE TABLE #t_CalculoLancamento
		     (PKId 	      INT)

	EXECUTE('INSERT INTO #t_CalculoLancamento
			SELECT PKid FROM tblReceita
			WHERE PKId IN('+@strPKId+')')

	--VARIÁVEIS A SEREM USADAS----------------------
	DECLARE @PKId		      AS INT,  	      --
		@intNumeroReceitas    AS INT,	      --
		@intContador	      AS INT,	      --
		@blnUsaFaixaDeValor   AS BIT,	      --	
		@blnECalculada 	      AS BIT,	      --
		@blnParcelar	      AS BIT,	      --
		@dblValor	      AS MONEY,	      --
		@dblValorAparcelar    AS MONEY,	      --			
		@dblIndexador	      AS MONEY,	      --	
		@dblExecFormula	      AS MONEY,	      --
		@strParametrosAUX     AS NVARCHAR(100)--
	--FIM DA DECLARAÇÃO ----------------------------

	--INICIALIZAÇÕES DE VARIÁVEIS-------------------------------------------------
	SET @dblIndexador = (SELECT ISNULL(IE.dblValor,0)			    --
			       FROM tblIndiceEconomico  IE, tblIndexadorEconomico E -- 
			      WHERE IE.intIndexador = E.PKId			    --
				AND IE.dtmData IN (SELECT MAX(IE.dtmData) 	    --
						     FROM tblIndiceEconomico IE))   --
	SET @strParametrosAUX = @strParametro					    --
	IF @bytUtilizacao = 2							    --
		SET @PKIdImobilEconomico = (SELECT PKId 			    --
					      FROM tblEconomico 		    --
					     WHERE strInscricaoCadastral = @strInscricaoCadastral)
	--FIM DA INICIALIZAÇÃO--------------------------------------------------------		
	SET @intNumeroReceitas = (SELECT COUNT(PKId) FROM #t_CalculoLancamento)
	SET @intContador = 0
	DECLARE	c_CalculoParcelaTaxa CURSOR FOR
		SELECT PKId FROM #t_CalculoLancamento
	OPEN	c_CalculoParcelaTaxa
	FETCH	c_CalculoParcelaTaxa INTO
		@PKId
	WHILE @@FETCH_STATUS = 0
	BEGIN
		SET @intContador = @intContador + 1
		SET @dblValorAparcelar = 0						    
		SELECT @blnUsaFaixaDeValor = blnUsaFaixaDeValor, @blnECalculada = blnECalculada,
		       @dblValor = dblValor, @blnParcelar = blnParcelar
		  FROM tblReceita 
		 WHERE PKId = @PKId  

		IF @blnECalculada = 1
		BEGIN
			EXECUTE sp_ParametroReceitas @PKId, @PKIdImobilEconomico,@strParametro OUTPUT
			EXECUTE sp_CalculoFormulaExecutada @PKId, @dblExecFormula OUTPUT, @strParametro
			SET @strParametro = @strParametrosAUX
			IF @blnUsaFaixaDeValor = 1 
			BEGIN
				SET @dblValor = (SELECT ISNULL(F.dblValor,0) 
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
					SET @dblValorAparcelar = ISNULL(@dblExecFormula,0) + 
							(ISNULL(@dblValor,0) *ISNULL(@dblIndexador,0))
			END
			ELSE
				SET @dblValorAparcelar = ISNULL(@dblExecFormula,0) 
		END
		ELSE
			SET @dblValorAparcelar =  ISNULL(@dblValor,0) 
		--Verifica Se Existe Isenção Para o Determinado Indivíduo
		EXECUTE sp_VerificaIsencao @bytUtilizacao, @strInscricaoCadastral, 
					   @intComposicaoReceita, @PKId, 
					   @dblValorAparcelar OUTPUT
		--------------------------------------------------------------------------------
		IF @dblValorAliquota != (-1) AND @blnParcelar = 0 -- EXISTE ALIQUOTA E PARCELÁVEL (ISSQN Estimado)
			SET @dblValorAparcelar = @dblValorAparcelar * @dblValorAliquota/100
		--------------------------------------------------------------------------------
		IF @intNumeroReceitas = @intContador
			SET @dblValorAparcelar = @dblValorDiferenca
		ELSE
			SET @dblValorDiferenca = @dblValorDiferenca - @dblValorAparcelar
		------------------------------------------------------------------------------
		EXECUTE sp_CalculoParcelaTaxaSub @dblValorAparcelar , @PKId, @intFlag,
			      	  		 @intParcelaInicial, @intParcelaFinal, 
						 @dtmDtParcela, @intExercicio, @intIntervalo, 
						 @gLngCodUsr, @dblValorDesconto 
		-------------------------------------------------------------------------------
		FETCH c_CalculoParcelaTaxa INTO
		      @PKId
	END
	CLOSE c_CalculoParcelaTaxa
	DEALLOCATE c_CalculoParcelaTaxa