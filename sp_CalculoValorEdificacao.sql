IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoValorEdificacao' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoValorEdificacao
GO

-- sp_CalculoValorEdificacao 4,0
CREATE PROCEDURE sp_CalculoValorEdificacao(@PKIdImobiliario AS INT,
		               		   @dblValor	    AS MONEY OUTPUT)
AS
	IF (SELECT ISNULL((SELECT bytEdificado FROM tblImobiliario
			   WHERE PKId = @PKIdImobiliario),0)) = 1
	BEGIN
		DECLARE @dblAreaConstruida  MONEY,
			@dblCAT		    MONEY,
			@dblFatoresCorrecao MONEY,
			@dblMT2deConstrucao MONEY,
			@strParametros	    NVARCHAR(100)

		SET @strParametros = ''+ CONVERT(NVARCHAR(30),@PKIdImobiliario) +', @dblValor OUTPUT'	
		--AREA CONTRUÍDA
		EXECUTE sp_CalculoFormulaExecutada -18, @dblAreaConstruida OUTPUT, @strParametros
		--CAT
		EXECUTE sp_CalculoFormulaExecutada -22, @dblCAT OUTPUT, @strParametros
		--FATORES DE CORREÇÃO
		EXECUTE sp_CalculoFormulaExecutada -23, @dblFatoresCorrecao OUTPUT, @strParametros
		--MT² DE CONSTRUÇÃO
		EXECUTE sp_CalculoFormulaExecutada -21, @dblMT2deConstrucao OUTPUT, @strParametros

		SET @dblValor = ISNULL(@dblAreaConstruida,0) * ISNULL(@dblCAT,0) 
				* ISNULL(@dblMT2deConstrucao,0) * ISNULL(@dblFatoresCorrecao,0)
	END
	ELSE
		SET @dblValor = (0.000)

