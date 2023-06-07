IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_VerificaIPTU' AND TYPE = 'P')
   DROP PROCEDURE sp_VerificaIPTU
GO 
-- sp_VerificaIPTU "02020020023001",10,2001,1,10 
CREATE PROCEDURE sp_VerificaIPTU(@strInscricaoCadastral AS NVARCHAR(100),
				 @dblDesconto		AS MONEY,
				 @intExercicio		AS INT,
				 @intParcelaInicial	AS INT,
				 @intParcelaFinal	AS INT)
AS
	DECLARE @PKId 	 	      INT,
		@dblIPU  	      MONEY,
		@dblITU  	      MONEY,
		@dblCCalcamento       MONEY,
		@dblIluminacaoPublica MONEY,
		@dblColetaDeLixo      MONEY,
		@dblTEXP 	      MONEY,
		@dblTestadaPrincipal  MONEY,
		@dblTestadaIdeal      MONEY,
		@dblCAT  	      MONEY,
		@dblFatoresDeCorrecao MONEY,
		@dblMT2DeContrucao    MONEY,
		@dblAreaConstruida    MONEY,
		@dblValorEdificacao   MONEY,
		@dblAreaTotalConst    MONEY,
		@dblAreaDoTerreno     MONEY,
		@dblFracaoIdeal	      MONEY,
		@dblMT2DoTerreno      MONEY,
		@dblTopografia	      MONEY,
		@dblSituacao	      MONEY,
		@dblPedologia	      MONEY,
		@dblValorVenalTerreno MONEY,
		@dblValorVenalImovel  MONEY,
		@strParametros	      NVARCHAR(100) --GUARDA PARÂMETROS DE SP'S

	SELECT @PKId = PKId
	  FROM tblImobiliario
	 WHERE strInscricaoAnterior  = @strInscricaoCadastral

	CREATE TABLE #t_TblTemporaria (strDescricao NVARCHAR(400),
					dblValor MONEY)
	CREATE TABLE #t_TblTemporaria2(dblValor MONEY,
					strDescricao NVARCHAR(400))
					

	SET @strParametros = ''+ CONVERT(NVARCHAR(30),@PKId) +', @dblValor OUTPUT'
	--IPU
	EXECUTE sp_CalculoFormulaExecutada 101, @dblIPU OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaDeCalculo
						WHERE intReceita = 101),@dblIPU
	--ITU
	EXECUTE sp_CalculoFormulaExecutada 97, @dblITU OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaDeCalculo
						WHERE intReceita = 97), @dblITU
	--CONSERVAÇÃO DE CALÇAMENTO
	EXECUTE sp_CalculoFormulaExecutada 68, @dblCCalcamento OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaDeCalculo
						WHERE intReceita = 68), @dblCCalcamento
	--ILUMINAÇÃO PÚBLICA
	EXECUTE sp_CalculoFormulaExecutada 67, @dblIluminacaoPublica OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaDeCalculo
						WHERE intReceita = 67), @dblIluminacaoPublica
	--COLETA DE LIXO
	EXECUTE sp_CalculoFormulaExecutada 65, @dblColetaDeLixo OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaDeCalculo
						WHERE intReceita = 65), @dblColetaDeLixo
	--CÁLCULO DA TESTADA PRINCIPAL
	EXECUTE sp_CalculoFormulaExecutada -20, @dblTestadaPrincipal OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 20), @dblTestadaPrincipal
	--CÁLCULO DA TESTADA IDEAL
	EXECUTE sp_CalculoFormulaExecutada -5, @dblTestadaIdeal OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 5), @dblTestadaIdeal
	--CÁLCULO DO CAT
	EXECUTE sp_CalculoFormulaExecutada -22, @dblCAT OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 22), @dblCAT
	--CÁLCULO DOS FATORES DE CORREÇÃO
	EXECUTE sp_CalculoFormulaExecutada -23, @dblFatoresDeCorrecao OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 23), @dblFatoresDeCorrecao
	--CÁLCULO DO MT² DE CONSTRUÇÃO
	EXECUTE sp_CalculoFormulaExecutada -21, @dblMT2DeContrucao OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 21), @dblMT2DeContrucao
	--AREA CONSTRUÍDA
	EXECUTE sp_CalculoFormulaExecutada -18, @dblAreaConstruida OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 18), @dblAreaConstruida
	--VALOR VENAL DA EDIFICAÇÃO
	EXECUTE sp_CalculoFormulaExecutada -2, @dblValorEdificacao OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 2), @dblValorEdificacao
	--AREA TOTAL CONTRUÍDA
	EXECUTE sp_CalculoFormulaExecutada -19, @dblAreaTotalConst OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 19), @dblAreaTotalConst
	--AREA DO TERRENO
	EXECUTE sp_CalculoFormulaExecutada -17, @dblAreaDoTerreno OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 17), @dblAreaDoTerreno
	--FRAÇÃO IDEAL
	EXECUTE sp_CalculoFormulaExecutada -4, @dblFracaoIdeal OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 4), @dblFracaoIdeal
	--MT² DO TERRENO
	EXECUTE sp_CalculoFormulaExecutada -16, @dblMT2DoTerreno OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 16), @dblMT2DoTerreno
	--CÁLCULO DA TOPOGRAFIA
	EXECUTE sp_CalculoFormulaExecutada -13, @dblTopografia OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 13), @dblTopografia
	--CÁLCULO SITUAÇÃO
	EXECUTE sp_CalculoFormulaExecutada -14, @dblSituacao OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 14), @dblSituacao
	--CÁLCULO DA PEDOLOGIA
	EXECUTE sp_CalculoFormulaExecutada -15, @dblPedologia OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 15), @dblPedologia
	--VALOR VENAL DO TERRENO
	EXECUTE sp_CalculoFormulaExecutada -1, @dblValorVenalTerreno OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 1), @dblValorVenalTerreno
	--VALOR VENAL DO IMÓVEL
	EXECUTE sp_CalculoFormulaExecutada -3, @dblValorVenalImovel OUTPUT, @strParametros
	INSERT INTO #t_TblTemporaria SELECT (SELECT CONVERT(NVARCHAR(4000),strDescricao) FROM tblFormulaBasica
						WHERE intCodigo = 3), @dblValorVenalImovel

	INSERT INTO #t_TblTemporaria2 
	SELECT dblValor AS dblValor, strDescricao AS strDescricao FROM #t_TblTemporaria


	BEGIN TRANSACTION Simulado	
		SET @strInscricaoCadastral = '"15,65,67,62,68,97,101","'+ @strInscricaoCadastral +'","' + @strInscricaoCadastral 
			+ '",16,'+ CONVERT(NVARCHAR,@intParcelaInicial)+ ',' + CONVERT(NVARCHAR,@intParcelaFinal)+ ',' 
			+ CONVERT(NVARCHAR(4),@intExercicio) +',9,' + CONVERT(NVARCHAR(10),@dblDesconto) 
			+ ',701' 
		EXECUTE sp_CalculoFormulaExecutada -24,NULL, @strInscricaoCadastral
		--Verificação
		/*
		DECLARE @MAXP_PKId INT
		SET @MAXP_PKId = (SELECT MAX(PKId) FROM tblLancamentoCalculo)
		IF ((SELECT SUM(dblValorParcela) FROM tblParcelaReceita
		     WHERE intNumeroParcela <> 0
		       AND intLancamentoCalculo = @MAXP_PKId) 
			= 
		   (SELECT SUM(dblValorParcela) FROM tblParcelaTaxa 
		     WHERE intNumeroParcela <> 0
		       AND intLancamentoCalculo = @MAXP_PKId)
			AND
			(SELECT SUM(dblValorParcela) FROM tblParcelaReceita
		     WHERE intNumeroParcela = 0
		       AND intLancamentoCalculo = @MAXP_PKId) 
			= 
		   (SELECT SUM(dblValorParcela) FROM tblParcelaTaxa 
		     WHERE intNumeroParcela = 0
		       AND intLancamentoCalculo = @MAXP_PKId)) 
			INSERT INTO #t_TblTemporaria2 				
			SELECT 0, ('Sucesso nos Lançamentos : Soma da Parcela Receita Igual a Soma da Parcela Taxa')
		ELSE
			INSERT INTO #t_TblTemporaria2 				
			SELECT -1, ('Falha nos Lançamentos : Soma da Parcela Receita Diferente da Soma da Parcela Taxa')
		*/
		--Parcela Receita
		INSERT INTO #t_TblTemporaria2 	
		SELECT (SELECT COUNT(*) FROM tblParcelaReceita WHERE lngCodUsr = 701) AS dblValor,
		' <<--NÚMERO DE INSERÇÕES NO PARCELAS RECEITAS-->>'

		INSERT INTO #t_TblTemporaria2 	
		SELECT dblValorParcela AS dblValor, 'Número = ' + CONVERT(NVARCHAR(10),intNumeroParcela)
	        + ' - Data de Vencimento = ' + CONVERT(NVARCHAR(20),DAY(dtmDataVencimento)) 
		+'/'+ CONVERT(NVARCHAR(20),MONTH(dtmDataVencimento)) 
		+'/'+ CONVERT(NVARCHAR(20),YEAR(dtmDataVencimento))  AS strDescricao
		 FROM tblParcelaReceita WHERE lngCodUsr = 701

		--Parcela Taxa
		INSERT INTO #t_TblTemporaria2 	
		SELECT (SELECT COUNT(*) FROM tblParcelaTaxa WHERE lngCodUsr = 701) AS dblValor,
		' <<--NÚMERO DE INSERÇÕES NO PARCELAS TAXAS-->>'

		INSERT INTO #t_TblTemporaria2 	

        	SELECT dblValorParcela AS dblValor, 'Número = ' + CONVERT(NVARCHAR(10),intNumeroParcela)
	        + ' - Data de Vencimento = ' + CONVERT(NVARCHAR(20),DAY(dtmDataVencimento)) 
		+'/'+ CONVERT(NVARCHAR(20),MONTH(dtmDataVencimento)) 
		+'/'+ CONVERT(NVARCHAR(20),YEAR(dtmDataVencimento)) 
		+ ' - Receita = ' + (SELECT strDescricao
				       FROM tblReceita
				      WHERE PKId = intReceita) AS strDescricao

		 FROM tblParcelaTaxa WHERE lngCodUsr = 701
	SELECT dblValor AS dblValor, strDescricao AS strDescricao FROM #t_TblTemporaria2
	ROLLBACK TRANSACTION Simulado

	DROP TABLE #t_TblTemporaria
	DROP TABLE #t_TblTemporaria2
	/*------------------------------------------------------------------
	SELECT ((-1)*intCodigo) AS IntCodigo, strNome FROM tblFormulaBasica
	UNION
	SELECT intReceita AS IntCodigo, strNome FROM tblFormulaDeCalculo
	------------------------------------------------------------------*/